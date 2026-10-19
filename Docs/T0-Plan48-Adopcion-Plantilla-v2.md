# T0 — Plan 48: adopción de la plantilla v2 y paridad con el manual julio-2

> **H0 read-only ejecutado 2026-10-19 por el `implementer`** tras la aprobación del Ingeniero
> (Clarification Gate §7 ratificado). **Nada se corrigió**: cero `<f>` creadas/modificadas/borradas,
> cero filas/columnas insertadas/borradas, cero escrituras sobre la v1 o la v2, cero commits.
>
> Fixture de la puerta: `Remuneracion.IntegrationTests/AdopcionPlantillaV2T0Plan48Tests.cs`
> → **21/21 verdes** (BCL `ZipArchive` + `XDocument`, sin Excel/COM ni writers/readers de producción).
>
> Entradas (SHA256 verificados antes y después):
>
> | Rol | Archivo | SHA256 |
> |---|---|---|
> | v1 histórica (no se toca) | `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC31…C5214` |
> | v2 nueva base (no se edita) | `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx` | `FF5CF7…2632` |
> | Manual v2 (nuevo oráculo) | `Docs/Prueba Julio-2/Resultado/Resultado Administrativo con plantilla estandar v2/Formato Remuneracion 202607-2 Total.xlsx` | (se lee) |
> | Corrida de referencia | `%TEMP%\opencode\p47-h2c\Remuneración 202607-2 Total.xlsx` | (se lee) |

---

## 0. Resumen ejecutivo

1. **v1 y v2 tienen las mismas 40 hojas, en el mismo orden y con los mismos nombres** (§1.1).
2. **El delta v1→v2 es pequeño y localizado**: R1 fila 649 (16 celdas → 0), R2 (bloques ASE2/ASE4
   +11 filas cada uno), `Valida - Control Recaudo` (E7/E10/F10 pasan de valor a fórmula), 5
   `VALIDACION_*` (C15=1, C14 sin caché) y las **fórmulas consumidoras** de R2 en
   `CONSOLIDADO_TOTAL RECAUDO` (61), `DetRetri2026092` (40), `DetValiRetri2026092` (51) y las 5
   `REMUNERACION_*` (60/56/60/60/60). **El resto de las 40 hojas no tiene delta** (§1.2, P6).
3. **P2 (nuestra corrida v1 vs manual v2):** de todas las diferencias, las que son **contenido real
   literal-contra-literal** son pocas y con causa citada (§2). La única brecha **interna** es **R2:
   9 filas / 53.283.379,41** que son **exactamente** las filas que `CerosDocumentadosDetalleR2R4`
   declara “sin slot” contra la v1 → **falso hueco probado**: la v2 sí tiene esos slots.
4. **P3 (delta del manual):** **492 fórmulas de R1** = 415 cambios de texto (337 `G→H`, 41 `F→G`,
   37 desplazamiento de fila en la fila 105) + **77 fórmulas nuevas** en 682/692/702 (+683/693/703).
   **La v2 NO trae esas 492**: el manual las introdujo al ejecutar. Decisión = P5(a) (§3, §5).
5. **P4 (factibilidad):** la App **sí puede** construir sobre la v2 sin ramas por período/empresa,
   pero **falla hoy** en la primera cadena congelada contra v1: la sonda real devolvió
   `ERR-PLANTILLA` accionable en `Rem. Anticipos R2!E144` (§4). Hay que **re-congelar** mapa de
   bloques R2, ceros documentados, cadenas de consumidores, control de recaudo (`D7:D10`/`F10`) y
   la cadena R1 682/702. **No hay STOP**: se absorbe con datos + re-congelamiento.
6. **Puerta H0→H1**: se puede abrir H1 **solo** sobre las hojas con delta probado (§8).

---

## 1. P1 — Matriz v1→v2 (bytes)

### 1.1 SHAs y las 40 hojas

- SHA256 v1 = `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` ✅
- SHA256 v2 = `FF5CF7A819110A1E1E2CA7BDAD553DD7A25141ADD472B1904BD496F082642632` ✅
- 40 hojas, **idéntico orden y nombres** en ambas (fixture `T0P48_P1_40Hojas_MismoOrdenYNombre`).

### 1.2 P6 — Tabla de paridad de las 40 hojas (v1→v2)

`fDiff` = texto de fórmula efectivo distinto (shared expandida; ausencia = cadena vacía).
`numDiff` = ambos numéricos y |Δ| > 0,5. Medido por el fixture (`T0P48_P6_*`).

| # | Hoja | fDiff | numDiff | Veredicto |
|---|---|---:|---:|---|
| 1 | `Recaudo EAAB Reciprocidad` | 0 | 0 | Sin delta (evidencia) |
| 2 | `Recaudo ENEL` | 0 | 0 | Sin delta (evidencia) |
| 3 | `Recaudo ENERBIT` | 0 | 0 | Sin delta (evidencia) |
| 4 | `Recaudo EAAB + Ciud Limp` | 0 | 0 | Sin delta (evidencia) |
| 5 | `Recaudo Directa Occidente` | 0 | 0 | Sin delta (evidencia) |
| 6 | `REPORTE RECAUDO x BANCO` | 0 | 0 | Sin delta (evidencia) |
| 7 | `ANT EXT-REV` | 0 | 0 | Sin delta (hoja protegida, no se escribe) |
| 8 | `ANTICIPOS USUARIOS` | 0 | 0 | Sin delta (evidencia) |
| 9 | `BCE SC POR FACT.` | 0 | 0 | Sin delta (evidencia) |
| 10 | `Reporte Componentes R1` | 0 | **16** | **Delta probado**: fila 649 → 0 |
| 11 | `Rem. Anticipos R2` | **352** | 0 | **Delta probado**: geometría nueva + OPORTUNO |
| 12 | `Reversion Pagos R4` | 0 | 0 | Sin delta (evidencia) |
| 13 | `SALDOS POR NOTA` | 0 | 0 | Sin delta (evidencia) |
| 14 | `RETRIBUCION NEGATIVA` | 0 | 0 | Sin delta (evidencia) |
| 15 | `AJUSTES - SF-T` | 0 | 0 | Sin delta (evidencia) |
| 16 | `INTERVENTORIA` | 0 | 0 | Sin delta (evidencia) |
| 17 | `Informe AFaseo Recaudo 202609-2` | 0 | 0 | Sin delta (evidencia) |
| 18 | `CONSOLIDADO_TOTAL RECAUDO` | **61** | 0 | **Delta de fórmulas**: consumidores R2 realineados |
| 19 | `VALIDACION_TOTAL` | 0 | **2** | **Delta probado**: C15 0→1, C14 sin caché |
| 20 | `GERENTES_TOTAL` | 0 | 0 | Sin delta (evidencia) |
| 21 | `REMUNERACION Reciprocidad EAAB` | **60** | 0 | **Delta de fórmulas**: consumidores R2 |
| 22 | `VALIDACION_RECIP` | 0 | **2** | **Delta probado**: C15/C14 |
| 23 | `GERENTES_Reciprocidad_ EAAB` | 0 | 0 | Sin delta (evidencia) |
| 24 | `REMUNERACION_ENEL` | **60** | 0 | **Delta de fórmulas**: consumidores R2 |
| 25 | `VALIDACION_ENEL` | 0 | **2** | **Delta probado**: C15/C14 |
| 26 | `GERENTES_ENEL` | 0 | 0 | Sin delta (evidencia) |
| 27 | `REMUNERACION_ENERBIT` | **60** | 0 | **Delta de fórmulas**: consumidores R2 |
| 28 | `VALIDACION_ENERBIT` | 0 | **2** | **Delta probado**: C15/C14 |
| 29 | `GERENTES_ENERBIT` | 0 | 0 | Sin delta (evidencia) |
| 30 | `REMUNERACION_OCCIDENTE_ Directa` | **60** | 0 | **Delta de fórmulas**: consumidores R2 |
| 31 | `VALIDACION_OCCIDENTE` | 0 | **2** | **Delta probado**: C15/C14 |
| 32 | `GERENTES_OCCIDENTE _ Directa` | 0 | 0 | Sin delta (evidencia) |
| 33 | `REMUNERACION_EAAB-CL` | **56** | 0 | **Delta de fórmulas**: consumidores R2 |
| 34 | `VALIDACION_EAAB-CL` | 0 | 0 | Sin delta (evidencia) |
| 35 | `GERENTES_EAAB-CL` | 0 | 0 | Sin delta (evidencia) |
| 36 | `DetRetri2026092` | **40** | 0 | **Delta de fórmulas**: consumidores R2 |
| 37 | `DetValiRetri2026092` | **51** | 0 | **Delta de fórmulas**: consumidores R2 |
| 38 | `Valida -Remunera` | 0 | 0 | Sin delta (evidencia) |
| 39 | `Valida - Anticipos` | 0 | 0 | Sin delta (evidencia) |
| 40 | `Valida - Control Recaudo` | **3** | **7** | **Delta probado**: E7/E10/F10 fórmulas |

> Nota metodológica (G-WORKBOOK/R-RUIDO): el diff textual crudo de valores es ruido de
> representación (`-335335691.51` vs `-335335691.50999999`) y los cachés `<v>` de celdas con `<f>`
> no son contenido escrito. Por eso la tabla usa `numDiff` con `cmpnum -Tol 0.5` y `fDiff` con
> *shared* expandida.

### 1.3 `Rem. Anticipos R2` — el cambio estructural (detalle)

**Bloques por ASE** (derivados por bytes: cabecera `B=<ASE>` + cierre `A=Total`):

| ASE | v1 (inicio-fin) | filas | v2 (inicio-fin) | filas | Δ |
|---|---|---:|---|---:|---:|
| 1 PROMOAMBIENTAL | 3–33 | 31 | 3–33 | 31 | 0 |
| 2 LIME | 67–113 | 47 | 67–124 | 58 | **+11** |
| 3 CIUDAD LIMPIA | 164–194 | 31 | 175–205 | 31 | 0 (shift +11) |
| 4 BOGOTA LIMPIA | 282–328 | 47 | 293–350 | 58 | **+11** (shift +11) |
| 5 AREA LIMPIA | 382–419 | 38 | 404–441 | 38 | 0 (shift +22) |

**Qué añaden exactamente ASE2 y ASE4** (misma lista de 11 filas en ambos, sin borrar ninguna):

```
C=3  D=Total            <- cierre de canal 3 (nuevo)
C=9  D=Total            <- cierre de canal 9 (nuevo)
D=Vlr Intereses         <- Vlr Intereses del canal 2 (antes sin slot)
D=Vlr Servicio          <- par VS del grupo H
D=Vlr Intereses         <- par VI del grupo H
D=Vlr Servicio          <- par VS del grupo O
D=Vlr Intereses         <- par VI del grupo O
C=3  D=Total            <- canal 3 en la mitad Subs/Cont
C=9  D=Total            <- canal 9 en la mitad Subs/Cont
D=Subsidio(-)/Contribucion(+)   <- Subs/Cont del canal 9
D=Subsidio(-)/Contribucion(+)   <- Subs/Cont del canal 9
```

ASE3 y ASE5 son **identidad estructural desplazada** (+11 y +22 respectivamente): ninguna fila
nueva, ningún borrado. ASE1 es identidad pura.

**Desplazamiento de los cuadros OPORTUNO** (fila `B=TOTAL C=OPORTUNO`):

| ASE | v1 | v2 |
|---|---:|---:|
| 1 | 47 | 47 |
| 2 | 144 | 155 |
| 3 | 262 | 273 |
| 4 | 362 | 384 |
| 5 | 446 | 468 |

**Cuadro OPORTUNO — fórmula vs escritura por fila** (v2; el cuadro **ya existía en v1** con las
mismas fórmulas; no es una adición de la v2). Por ASE, filas `(empresa, ¿E es fórmula?, fila
Subsidio)`:

| ASE | TOTAL | RECIPROCIDAD | ENEL | ENERBIT | OCCIDENTE | CIUDAD LIMPIA-ACUEDUCTO |
|---|---|---|---|---|---|---|
| 1 | 47 fórm. / 48 | 50 **literal 0** / 51 | 53 fórm. / 54 | 56 **literal 0** / 57 | 59 fórm. / 60 | 62 **literal 0** / 63 |
| 2 | 155 fórm. / 156 | 158 **fórm.** / 159 | 161 fórm. / 162 | 164 **literal 0** / 165 | 167 fórm. / 168 | 170 **literal 0** / 171 |
| 3 | 273 fórm. / 274 | 276 **literal 0** / 277 | 279 fórm. / 280 | 282 **literal 0** / 283 | 285 fórm. / 286 | 288 **literal 0** / 289 |
| 4 | 384 fórm. / 385 | 387 **fórm.** / 388 | 390 fórm. / 391 | 393 **literal 0** / 394 | 396 fórm. / 397 | 399 **literal 0** / 400 |
| 5 | 468 fórm. / 469 | 471 **literal 0** / 472 | 474 fórm. / 475 | 477 **fórm.** / 478 | 480 fórm. / 481 | 483 **literal 0** / 484 |

**Respuesta empírica a P5(b):** ninguna fila del cuadro se “llena con valores” por la App. Las
filas con `<f>` las calcula Excel (el writer debe **no** tocarlas → guard anti-fórmula); las filas
sin `<f>` son **literal 0** de empresas que no existen en ese ASE y el manual también las deja en 0.
**No hay fila escribible en el cuadro OPORTUNO.**

Fórmulas testigo del TOTAL (v2): `E47=E21+E32-K21`, `E155=E101+E123-K101`,
`E273=E193+E204-K193`, `E384=E327+E349-K327`, `E468=E440+E426-K426`.

**Reconciliación del “+1.323 celdas nuevas”** (el plan lo citaba de `cmpnum`):
- `cmpnum onlyB = 1.323` = celdas con **contenido** en v2 sin contenido en v1 (cuenta celdas
  presentes-vacías en v1 como “nuevas”).
- `diffv onlyB = 312` = celdas con valor o fórmula **no vacíos** y **ausentes** del XML de v1.
- Medición propia: **793 refs** presentes en el XML de v2 y ausentes del XML de v1.
Las tres cifras son correctas con definiciones distintas; la que describe “celdas que ganaron
contenido” es **1.323**.

### 1.4 `Reporte Componentes R1` — fila 649 y zona 682/702

**Fila 649 (16 celdas, todas columnas F–T y Z):** la v1 trae **valores residuales**; la v2 los deja
en **0** (sin fórmula):

| Celda | v1 | v2 |
|---|---:|---:|
| `F649` | 13.359.789.539,23 | 0 |
| `G649` | 697.483.550,35 | 0 |
| `H649` | 388.259.624,73 | 0 |
| `I649` | 206.453.897,29 | 0 |
| `J649` | 3.256.933.254,80 | 0 |
| `K649` | 106.089.721,18 | 0 |
| `L649` | 12.154.250,72 | 0 |
| `M649` | 930.507.799,59 | 0 |
| `N649` | 903.844.784,82 | 0 |
| `O649` | 4.612.204.182,57 | 0 |
| `P649` | 2.072.913.975,94 | 0 |
| `Q649` | 173.121.721,61 | 0 |
| `R649` | −177.200,85 | 0 |
| `S649` | 191,97 | 0 |
| `T649` | 6,64 | 0 |
| `Z649` | −222,13 | 0 |

La fila 649 es `A=Componente B=Total` (interior del bloque ASE5). **El manual v2 la vuelve a
escribir** con los valores del período (`F649 = 13.084.651.457,06`), y **nuestra corrida `p47-h2c`
coincide con el manual**. Conclusión: no es un hueco; el mirror la escribe. La v2 solo limpia el
residuo histórico de la plantilla.

**Zona 682/702:** en la v2 estas filas están **vacías (sin fórmula)**; el manual les añade fórmulas
al ejecutar (ver P3): `F682=F660+F629-L629`, `F702=F659+F628-L628`, `F692=F652+F618-L618`, más los
bloques shared `H682:AO682` (=`H629`) y `H702:AO702` (=`H628`) y las celdas `F683/F693/F703`.
**No son de la v2**; son delta del manual (P5a).

### 1.5 `Valida - Control Recaudo`

| Celda | v1 | v2 | Manual |
|---|---|---|---|
| `E7` | valor 68.245.142.592 | **fórmula `SUM(D7:D9)`** (cache 0) | fórmula |
| `E10` | valor 5.736.574.007 | **fórmula `SUM(D10)`** (cache 0) | fórmula |
| `F10` | valor 73.981.716.599 | **fórmula `E7+E10`** (cache 0) | fórmula |
| `D7:D10` | valores (heredados de la plantilla) | **0** | 68.144.618.192 / 28.645.323 / 40.117.400 / 4.284.569.052 |

El manual deja `F10 = 72.497.949.967`, **idéntico al literal que escribe nuestra corrida**; la
diferencia es que la v2 **calcula** E7/E10/F10 desde `D7:D10`, y la App **hoy no escribe `D7:D10`**
(escribe `F10` como literal). Ver §4.

### 1.6 `VALIDACION_*` (ENEL/ENERBIT/OCCIDENTE/RECIP/TOTAL)

En las 5: **`C15: 0 → 1`** (prefijado) y **`C14` sin caché** (0). `VALIDACION_EAAB-CL` no cambia.
Es un cambio de caché/prefijado del **gate**, no de contenido de negocio (Plan 40: se evalúa).

### 1.7 Consumidores de R2 ya realineados en la v2

`CONSOLIDADO_TOTAL RECAUDO!D28:D32` apunta a las filas nuevas:

| Celda | v1 | v2 |
|---|---|---|
| `D28` | `'Rem. Anticipos R2'!E47` | `'Rem. Anticipos R2'!E47` |
| `D29` | `'Rem. Anticipos R2'!E144` | `'Rem. Anticipos R2'!E155` |
| `D30` | `'Rem. Anticipos R2'!E262` | `'Rem. Anticipos R2'!E273` |
| `D31` | `'Rem. Anticipos R2'!E362` | `'Rem. Anticipos R2'!E384` |
| `D32` | `'Rem. Anticipos R2'!E446` | `'Rem. Anticipos R2'!E468` |

Muestras testigo: `H29` `I144→I155`, `G31` `H362→H384`, `AE32` `E447→E469`.
El mismo realineamiento explica los `fDiff` de `DetRetri2026092` (40), `DetValiRetri2026092` (51) y
las 5 `REMUNERACION_*` (60/56/60/60/60).

---

## 2. P2 — Clasificación nuestra corrida (`p47-h2c`, v1) vs manual v2

Metodología triple: `cmpnum -Tol 0.5` (numérico), `diffv` (fórmulas con *shared* expandida) y
`cmplit` (**literales-contra-literales**: solo celdas sin `<f>` en ambos). Los cachés de fórmula
(`fCacheDiff`) **no son contenido**.

### 2.1 Tabla por hoja (solo hojas con diferencias)

| Hoja | litVsLit | fTextDiff | fCacheDiff | clase | onlyA | onlyB | Clasificación |
|---|---:|---:|---:|---:|---:|---:|---|
| `Reporte Componentes R1` | **0** | 415 | 464 | 77 | 0 | 0 | Solo fórmulas (delta del manual) → **contenido literal idéntico** |
| `Rem. Anticipos R2` | 2.960 | 2 | 41 | 350 | 237 | 312 | **Geometría v2** + falso hueco (ver §2.2) |
| `Reversion Pagos R4` | **0** | 0 | 224 | 0 | 0 | 0 | Solo cachés → **contenido idéntico** |
| `SALDOS POR NOTA` | **0** | 0 | 49 | 0 | 0 | 0 | Solo cachés → **contenido idéntico** |
| `REPORTE RECAUDO x BANCO` | **0** | 0 | 70 | 0 | 0 | 0 | Solo cachés |
| `AJUSTES - SF-T` | **0** | 0 | 203 | 0 | 0 | 0 | Solo cachés |
| `VALIDA - Control Recaudo` | **7** | 0 | 0 | 0 | 0 | 0 | **Contenido real**: D7:D10 (ver §2.3) |
| `BCE SC POR FACT.` | **2** | 0 | 39 | 0 | 0 | 0 | **Contenido real**: `E7`/`D7` (divergencia V5 declarada) |
| `CONSOLIDADO_TOTAL RECAUDO` | **1** | 60 | 834 | 2 | 2 | 0 | Fórmulas (consumidores) + `D6` metadato |
| `DetRetri2026072` | 59 | 40 | 72 | 0 | 0 | 0 | Columnas **excluidas por diseño** + fórmulas |
| `DetValiRetri2026072` | 63 | 51 | 69 | 0 | 0 | 0 | Columnas **excluidas por diseño** + fórmulas |
| `Recaudo ENEL` | 48 | 0 | 0 | 0 | 0 | 0 | **Q1 acumulada** por diseño (V3) |
| `Recaudo Directa Occidente` | 48 | 0 | 0 | 0 | 0 | 0 | idem |
| `Recaudo EAAB Reciprocidad` | 18 | 0 | 0 | 0 | 0 | 0 | idem |
| `Recaudo ENERBIT` | 16 | 0 | 0 | 0 | 0 | 0 | idem |
| `ANT EXT-REV` | 34 | 0 | 71 | 0 | 0 | 0 | Hoja **protegida** (la App no la escribe) |
| `REMUNERACION_*` (5) | **0** | 60/56/60/60/60 | 373/32/774/163/682 | 0 | 0 | 0 | Solo fórmulas (consumidores) |
| `GERENTES_*` (6) | **0** | 0 | N | 0 | 0 | 0 | Solo cachés |
| `VALIDACION_*` (5) | **0** | 0 | N | 0 | 0 | 0 | Solo cachés |
| `Valida - Anticipos` / `Valida -Remunera` | **0** | 0 | 124/488 | 0 | 0 | 0 | Solo cachés |
| `Informe AFaseo Recaudo 202607-2` | **0** | 0 | 69 | 0 | 0 | 0 | Solo cachés |
| `ANTICIPOS USUARIOS` | **0** | 0 | 53 | 0 | 0 | 0 | Solo cachés |

### 2.2 La única brecha interna: R2 es EXACTAMENTE el falso hueco de `CerosDocumentados`

Comparación por **suma de columnas por bloque ASE** (geometrías alineadas; ASE1/3/5 dan Δ≈0):

| ASE | ΔE (manual − corrida) | Causa |
|---|---:|---|
| 1 | 0,00 | Paridad total |
| 2 | **1.664.245,00** | 406,42 + 831.919,29 + 831.919,29 = las 3 filas documentadas “sin slot” |
| 3 | 0,00 | Identidad (shift +11) |
| 4 | **51.619.134,41** | −8.785,59 + 20.152.773,11 − 2.377,86 + 20.150.395,25 + 5.663.564,75 + 5.663.564,75 = las 6 filas documentadas “sin slot” |
| 5 | 0,00 | Identidad (shift +22) |
| **Total** | **53.283.379,41** | = las **9 filas de julio-2** de `CerosDocumentadosDetalleR2R4` |

Prueba en bytes: la v1 **no tiene** `C=3` ni `C=9` en ASE2/ASE4; la v2 **sí** (`C=9` en ASE2 R89 y
ASE4 R315, `C=3` en ASE2 R83 y ASE4 R309). El manual llena esos slots (`E87/E89 = 831.919,29`;
`E313 = 20.152.773,11`, `E314 = −2.377,86`, `E315 = 20.150.395,25`).
→ **Las 12 claves / 9 filas de julio-2 documentadas en el Plan 47 son un falso hueco bajo la v2.**
H1 debe **re-congelar** `CerosDocumentadosDetalleR2R4` (retirar las claves que ahora tienen slot) y
`BloquesR2` (rangos v2). El `Cantidad` de las claves que también ocurrían en agosto-2 debe
recalcularse contra la geometría v2.

### 2.3 Contenido real fuera de R2 (lista cerrada, con causa)

| Hoja / celdas | Diferencia | Clasificación |
|---|---|---|
| `Valida - Control Recaudo!D7:D10` | La corrida v1 conserva los valores de la plantilla v1; el manual escribe los totales de `Recaudo *!F27` (68.144.618.192 / 28.645.323 / 40.117.400 / 4.284.569.052) | **Brecha de adopción**: en v2 `D7:D10 = 0` y E7/E10/F10 son fórmulas → la App debe escribir `D7:D10` (§4) |
| `BCE SC POR FACT.!E7/D7` | −1.266.769.543,80 / −2.477.195.172,11 | **Divergencia declarada V5** (ASE5, dueño UAESP; Plan 29) — no es regresión |
| `CONSOLIDADO_TOTAL RECAUDO!D6` | 46.238 → 46.297 (fecha de proceso) | **Metadato**: la App no duplica `D6` (Plan 28 R-S-3); el manual tecleó otra fecha |
| `Recaudo *` D/E/I/J | 0 en la corrida Q2; el manual acumula Q1 y TOTAL MES | **Por diseño V3** (cada quincena escribe solo sus columnas) |
| `DetRetri`/`DetValiRetri` E/F/G/H/I/K/M/N/O y D..N de DetValiRetri | 0 en la corrida | **Excluidas por diseño** (`WorkbookLeafCellMapDetRetri.ColumnasExcluidas`, motivo `MotivoCadenaFormula`); `D`/`J`/`L`/`I`/`O` sí se escriben |
| `ANT EXT-REV` (34 celdas) | 0 en la corrida | **Hoja protegida**: la App valida sus fórmulas, no la escribe |
| `Reporte Componentes R1` | litVsLit = **0** | **Paridad de contenido**; solo cambian fórmulas (delta del manual) |

**Conclusión P2:** no hay una brecha de contenido “silenciosa”. La única brecha interna es el falso
hueco de R2 (53.283.379,41); el resto es geometría/caché/diseño/declarada/metadato, cada una con
causa citada.

### 2.4 Medición vs insumo (R2, `RerpoteDetalleSaldosaFavor`)

Las 9 filas del falso hueco **existen en la fuente** con exactamente esos totales:

| ASE | Fila fuente | Etiqueta | E (Total) | Celda v2 donde el manual la escribe |
|---|---:|---|---:|---|
| 2 | R13 | `Vlr Intereses` (canal 2) | 406,42 | grupo canal 2 |
| 2 | R18 | `Vlr Servicio` (canal 9) | 831.919,29 | `E87` |
| 2 | R19 | `C=9 D=Total` | 831.919,29 | `E89` |
| 4 | R13 | `Vlr Intereses` (canal 2) | −8.785,59 | grupo canal 2 |
| 4 | R18 | `Vlr Servicio` (canal 9) | 20.152.773,11 | `E313` |
| 4 | R19 | `Vlr Intereses` (canal 9) | −2.377,86 | `E314` |
| 4 | R20 | `C=9 D=Total` | 20.150.395,25 | `E315` |
| 4 | R42 | `Subsidio(-)/Contribucion(+)` (canal 9) | 5.663.564,75 | mitad Subs/Cont |
| 4 | R43 | `C=9 D=Total` (Subs/Cont) | 5.663.564,75 | mitad Subs/Cont |

Es decir: **fuente → manual** cierra al centavo, y **nuestra corrida** queda corta exactamente por
esas filas. No es un artefacto de comparación: es una omisión real que la v2 permite corregir.

---

## 3. P3 — Inventario del delta del manual (v2 → manual)

### 3.1 Las 492 fórmulas de `Reporte Componentes R1`

`fDiff(v2, manual) = 492` = **415 cambios de texto** + **77 fórmulas nuevas**.

| Patrón | Celdas | Detalle |
|---|---:|---|
| `G→H` exacto | **337** | p. ej. `P570: G486→H486`, `J95: G46+G15→H46+H15`, `I264: G150+G112→H150+H112` |
| `F→G` exacto | **41** | p. ej. `T284: F120→G120`, `K445: F299→G299` |
| Desplazamiento de fila en fila 105 | **37** | 36 celdas `F17→G16` + `F105: F17→F16` |
| **Subtotal cambios** | **415** | |
| Fórmulas nuevas fila 682 | **36** | `F682=F660+F629-L629`; `G682=G629`; master shared `H682:AO682` = `H629` |
| Fórmulas nuevas fila 692 | **2** | `F692=F652+F618-L618`; `G692=G618` |
| Fórmulas nuevas fila 702 | **36** | `F702=F659+F628-L628`; `G702=G628`; master shared `H702:AO702` = `H628` |
| Fórmulas nuevas 683/693/703 | **3** | `F683=F660`; `F693=F652`; `F703=F659` |
| **Subtotal nuevas** | **77** | |
| **TOTAL** | **492** | |

- Las 36/36 de 682/702 incluyen los *followers* shared `H..AO`; el volcado del T0 muestra el texto
  del **master** (`H629`/`H628`) porque el desplazamiento relativo de una *shared formula* es
  responsabilidad de Excel, no del XML.
- `W682`, `S682`, `AN682` (y el resto `H..AO`) **no** son fórmula en la v2; el manual las añade.
- La v2 **no trae** ninguna de las 492. **Decisión = P5(a)**: ¿la referencia oficial de R1 es la v2
  tal cual o la v2 + estas correcciones?

### 3.2 Literales del manual por hoja (v2 → manual, `litVsLit`)

| Hoja | litVsLit | Naturaleza |
|---|---:|---|
| `Reporte Componentes R1` | 3.435 | Escritura normal del período (la App también la hace) |
| `Rem. Anticipos R2` | 2.099 | Escritura normal + slots nuevos |
| `Reversion Pagos R4` | 824 | Escritura normal |
| `SALDOS POR NOTA` | 233 | Escritura normal |
| `Recaudo *` | 84/36/84/24 | Pegado real + Q1 acumulada |
| `ANT EXT-REV` | 34 | Hoja externa/protegida (manual) |
| `Recaudo *` / `BCE` / `Valida-Control` | 10 / 6 / 6 | Balance + control |
| `DetRetri*`/`DetValiRetri*` | — | Hojas renombradas por el preparador (2026072) |

La comparación **contra nuestra corrida** (§2) demuestra que la App ya escribe esos literales:
`R1 litVsLit = 0`, `R4 = 0`, `SALDOS = 0`, `banco = 0`, `AJUSTES = 0`, `REMUNERACION_* = 0`,
`GERENTES_* = 0`. **El delta del manual no es una brecha de contenido salvo el falso hueco R2.**

---

## 4. P4 — Veredicto de factibilidad de adopción

### 4.1 Sonda real del pipeline con la v2 (solo a temp; la v2 no se tocó)

```
dotnet run --project Remuneracion.Cli --no-build -- \
  --periodo 2026072 --carpeta "Docs/Prueba Julio-2/Insumos" \
  --plantilla "Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx" \
  --salida "%TEMP%/opencode/p48-h0-v2probe"
```

Resultado: **`ERR-PLANTILLA` (código 2)** — fail-fast accionable:

```
La celda 'E144' de 'Rem. Anticipos R2' debería seguir siendo fórmula; se detectó un valor fijo.
```

Causa: `WorkbookLeafCellMapQ2` / `WorkbookLeafCellMapPorAse` declaran la cadena de R2 con
`E144 = E94+E112-K94` (v1). En v2 esa celda está vacía y la misma celda lógica se desplazó al
cuadro OPORTUNO de ASE2 (`E155 = E101+E123-K101`). **El guard hace su trabajo** (no hubo escritura
parcial). SHA de la v2 verificado tras la sonda: intacto.

### 4.2 Qué falla y dónde (lista cerrada)

| # | Punto congelado contra v1 | Falla en v2 | Evidencia |
|---|---|---|---|
| 1 | `WorkbookLeafCellMapDetalleR2R4.BloquesR2` (3-33, 67-113, 164-194, 282-328, 382-419) | Rangos v2: 3-33, 67-124, 175-205, 293-350, 404-441 | §1.3 |
| 2 | `WorkbookLeafCellMapQ2`/`PorAse` cadena R2 (`E144`…) | `E144` sin fórmula; la cadena vive en `E155/E273/E384/E468` | §4.1 |
| 3 | `CerosDocumentadosDetalleR2R4` (canal 9, canal 2 VI de ASE2/ASE4) | La v2 sí tiene slot → **falso hueco**; no se puede seguir documentando el 0 | §2.2 |
| 4 | `CONSOLIDADO_TOTAL RECAUDO` esperado `E47/E144/E262/E362/E446` | v2 usa `E47/E155/E273/E384/E468` (ya realineado en la plantilla) | §1.7 |
| 5 | Cadenas `DetRetri2026092`/`DetValiRetri2026092`/`REMUNERACION_*` | v2 ya apunta a las filas nuevas (40/51/60 fórmulas) | §1.7 |
| 6 | `EscribirControlRecaudoTotal` escribe `Valida - Control Recaudo!F10` como literal | En v2 `F10` es **fórmula** `E7+E10` → guard anti-fórmula | §1.5 |
| 7 | La App no escribe `Valida - Control Recaudo!D7:D10` | En v2 `D7:D10 = 0` y `E7/E10/F10` son fórmulas → sin `D7:D10` la cadena da 0 y `C15` falla | §1.5 |
| 8 | Cadena R1 (`F75/F77`) y cadenas de control `682/702` | La v2 deja `682/702` sin fórmula; el manual las añade | §1.4, §3.1 |
| 9 | `ValidarCadenaFormulaR1R2R4` (`E47`, `D73`, `F75/F77`) | `E47` y `D73` no cambian; sí cambian los consumidores de R2 | §1.3, §1.7 |

### 4.3 Re-congelamientos necesarios (H1)

1. **`WorkbookLeafCellMapDetalleR2R4.BloquesR2`** → rangos v2 (datos, no código).
2. **`CerosDocumentadosDetalleR2R4`** → retirar las claves de ASE2/ASE4 que ya tienen slot en v2
   (canal 9 y canal 2 VI); recalcular `Cantidad` de las que compartían agosto-2.
3. **`WorkbookLeafCellMapQ2`/`WorkbookLeafCellMapPorAse`** → cadena R2 a `E47/E155/E273/E384/E468`.
4. **`CONSOLIDADO`/`DetRetri`/`DetValiRetri`/`REMUNERACION_*`** → verificar que las fórmulas de la
   v2 (ya realineadas) son las que el gate espera; re-congelar los `fragmentos` de validación.
5. **Control de recaudo** → escribir `D7:D10` desde `Recaudo *!F27` (por etiqueta) y **dejar de
   escribir `F10`** cuando sea fórmula (o escribir solo los destinos sin `<f>`); guard intacto.
6. **R1 682/702** → decisión P5(a); si se adoptan, es una excepción `<f>` autorizada, no un parche.
7. **R1 fila 649** → verificar que el mirror la escribe (hoy lo hace: corrida = manual); no requiere
   cambio de mapa, solo un assert de no-regresión.

**¿Se absorbe sin ramas por período/empresa? SÍ.** Todo lo anterior son **datos geométricos** y
**cadenas de referencia**, no condicionales por período/empresa (D-H intacto). **No hay STOP.**

---

## 5. P5 — Preguntas para la propietaria (texto para trasladar)

1. **(a) Referencia oficial de R1:** al ejecutar julio-2 usted cambió **492 fórmulas** de
   `Reporte Componentes R1` (337 `G→H`, 41 `F→G`, 37 en la fila 105, y 77 nuevas en 682/692/702).
   **La plantilla v2 `vf` no trae esas 492 fórmulas.** ¿La referencia oficial es la **v2 tal cual**
   o la **v2 + esas correcciones del manual**?
2. **(b) Cuadro OPORTUNO de R2:** sus filas son **fórmula** (TOTAL/Subsidio/ENEL/OCCIDENTE, y
   RECIPROCIDAD en ASE2/ASE4) o **literal 0** (empresas ausentes). ¿Confirmа que **ninguna** fila se
   llena con valores por la App y que las de empresa ausente quedan en 0?
3. **(c) Base desde ahora:** ¿la **v2** es la base para **todos** los períodos desde ahora y la v1
   queda como histórica?
4. **(d) `R4-Balance…`:** el archivo `R4-BalanceSubsidioyContribuciones_…xlsx` que está en la misma
   carpeta del manual, ¿es **fuente** del run o **anexo** de evidencia?

---

## 6. P6 — Veredicto por hoja (delta probado vs sin delta con evidencia)

La tabla completa está en §1.2. Resumen:

- **Delta probado (candidatas a H1):**
  `Reporte Componentes R1` (649 + zona 682/702), `Rem. Anticipos R2` (bloques + OPORTUNO + falso
  hueco), `Valida - Control Recaudo` (E7/E10/F10 + D7:D10), `VALIDACION_*` (5; C15/C14).
- **Delta de fórmulas consumidoras (arrastre ya realineado en v2; verificar en H1):**
  `CONSOLIDADO_TOTAL RECAUDO`, `DetRetri2026092`, `DetValiRetri2026092`, `REMUNERACION_*` (5).
- **Sin delta con evidencia (no se tocan):** las 24 hojas restantes (Recaudo *, banco, ANT EXT-REV,
  ANTICIPOS USUARIOS, BCE, R4, SALDOS, RETRIBUCION NEGATIVA, AJUSTES, INTERVENTORIA, Informe AFaseo,
  GERENTES_* (6), VALIDACION_EAAB-CL, Valida -Remunera, Valida - Anticipos).

---

## 7. Gates del H0

| Gate | Resultado |
|---|---|
| **G-BUILD** | `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores** |
| **G-SUITE** | Ver §7.1 (trx `p48-suite.trx`): set-diff vs `p47-warnings.trx` |
| **G-Q1** | Q1 no se tocó (H0 read-only); los tests Q1 siguen en el set base |
| **G-SHA** | v1 `3086AC31…C5214` y v2 `FF5CF7…2632` intactos antes y después (fixture + sonda) |
| **G-PROC** | **cero commits**; cero código de producción; solo se agregó el fixture read-only |
| **Fixture** | `AdopcionPlantillaV2T0Plan48Tests` → **21/21 verdes**, 0 Skip |

### 7.1 Suite

- Corrida observada (`p48-suite.trx`): **1058 Total / 1020 Passed / 38 Failed / 0 Skipped**
  (base `p47-warnings.trx`: 999 Passed / 1037 Total; los 21 tests nuevos del fixture entran en
  Passed).
- **Set-diff de fallos vs `o43-final-failed.txt` (39): 0 altas / 1 baja.** La baja es
  `Remuneracion.IntegrationTests.ProcesadorPeriodoTests.Ejecutar_Periodo2026072_ConLectorOracle_VeredictosValidacionesNoVacios`
  (ROJO→VERDE), **heredada del Plan 46** y ya reconciliada en `p47-warnings-setdiff.txt`: el test
  corre Q2 con oráculo y sin reconocimiento; pre-fix la corrida bloqueaba con la divergencia
  declarada OCCIDENTE-ASE5 y no producía veredictos; con el espejo R1 consistente cierra y verifica
  los 5 bloques VALIDACIONES (HU-13/2.7). **Plan 48 no introduce altas ni bajas nuevas.**
- **G-Q1:** `Remuneracion.IntegrationTests.SeccionEspejoR1Tests.T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback` → **Passed**.
- **Fixture nuevo:** `AdopcionPlantillaV2T0Plan48Tests` → **21 tests / 0 fallos**, 0 Skip.

> **Nota de entorno:** el fixture nuevo compara la corrida de referencia
> `%TEMP%\opencode\p47-h2c\Remuneración 202607-2 Total.xlsx` (S1.6). Si esa corrida no existe, el
> test falla con mensaje accionable (dependencia declarada, igual que los insumos del fixture 47).

---

## 8. Veredicto de la puerta H0→H1

**(a) Hojas con delta probado (candidatas a H1):**

1. **`Rem. Anticipos R2`** — re-congelar `BloquesR2` (v2), `CerosDocumentadosDetalleR2R4`
   (falso hueco), cadena `E155/E273/E384/E468` y consumidores. Es la única brecha de contenido
   (53.283.379,41).
2. **`Reporte Componentes R1`** — fila 649 (limpieza de plantilla; el mirror la escribe) y zona
   682/702 (**decisión P5a**; si se adopta, excepción `<f>` autorizada).
3. **`Valida - Control Recaudo`** — E7/E10/F10 fórmulas en v2 + escribir `D7:D10` (y no `F10`).
4. **`VALIDACION_*` (5)** — C15 prefijado / C14 sin caché; el gate debe evaluar, no pegar.

**(b) Hojas sin delta con evidencia (no se tocan):** las 24 restantes de §6, con `fDiff=0` y
`numDiff=0` probados por el fixture.

**No hay STOP:** la geometría v2 se absorbe con datos y re-congelamiento, sin ramas por
período/empresa. **La puerta H0→H1 queda documentada; la cierra el Ingeniero.**

---

### Anexo A — Herramientas y comandos (read-only, `%TEMP%\opencode\`)

- `dumpxlsx.ps1 -Path -Sheet -From -To -Cols -List -OnlyValues`
- `diffv.ps1 -A -B -Sheet -Max -Summary -NoValues` (fórmulas con shared expandida)
- `cmpnum.ps1 -A -B -Sheet -Max -Summary -Tol 0.5` (comparador numérico)
- `cmplit.ps1 -A -B -Sheet -Max -Summary` (literales-contra-literales, aislado para este T0)
- `sumrange.ps1 -A -B -Sheet -RangesA -RangesB` (suma por columna/bloque, geometrías alineadas)

> Advertencias aplicadas: (i) el diff textual crudo es ruido de representación → `cmpnum -Tol 0.5`;
> (ii) los cachés `<v>` de celdas con `<f>` no son contenido escrito; (iii) el diff de fórmulas sin
> expandir *shared* es ruido.

---

## 9. H1/H2 — evidencia de cierre (2026-10-19)

> H1 autorizado por el Ingeniero **con el alcance probado y SIN adoptar las 492 fórmulas del manual**.
> Detalle de cambios y trazabilidad: `plans/48` §9.2.

### 9.1 Corrida fresca sobre la v2

```
CLI --periodo 2026072 --carpeta "Docs/Prueba Julio-2/Insumos" \
    --plantilla "Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx" \
    --salida "%TEMP%/opencode/p48-h2-fresh"
→ RESULTADO OK codigo=0
```

`VALIDACION_TOTAL!C14 = 72.497.949.967` (sellado por read-back) y `C15 = 1`: el gate honesto (W-1/E1-1)
cierra sobre la v2.

### 9.2 Tabla corrida fresca → manual v2 (`cmplit`: literales-contra-literales)

| Hoja | litVsLit | fTextDiff | fCacheDiff | clase | onlyA | onlyB | Clasificación |
|---|---:|---:|---:|---:|---:|---:|---|
| `Rem. Anticipos R2` | **0** | 0 | 238 | 0 | 0 | 0 | **Contenido idéntico** (238 = cachés de fórmulas) |
| `Reporte Componentes R1` | **0** | 415 | 464 | 77 | 0 | 0 | Solo fórmulas: **492 del manual** (NO adoptadas) |
| `Reversion Pagos R4` | **0** | 0 | 224 | 0 | 0 | 0 | Contenido idéntico (cachés) |
| `SALDOS POR NOTA` | **0** | 0 | 49 | 0 | 0 | 0 | Contenido idéntico |
| `REPORTE RECAUDO x BANCO` | **0** | 0 | 70 | 0 | 0 | 0 | Contenido idéntico |
| `AJUSTES - SF-T` | **0** | 0 | 203 | 0 | 0 | 0 | Contenido idéntico |
| `REMUNERACION_*` (5) | **0** | 60/56/60/60/60 | 373/32/774/163/682 | 0 | 0 | 0 | Solo fórmulas (consumidores) |
| `GERENTES_*` (6) | **0** | 0 | N | 0 | 0 | 0 | Contenido idéntico |
| `VALIDACION_*` | **0** | 0 (EAAB-CL 5, ENERBIT 4, RECIP 2) | N | 0 | 0 | 0 | Cachés + tokens de período (D↔F) |
| `Valida - Control Recaudo` | 1 (`F9`) | 0 | 0 | 3 (`E7/E10/F10`) | 0 | 0 | Representación: App conserva fórmulas v2; manual puso valores + `F9=1` |
| `BCE SC POR FACT.` | 2 (`E7/D7`) | 0 | 39 | 0 | 0 | 0 | **Divergencia declarada V5** (ASE5) |
| `CONSOLIDADO_TOTAL RECAUDO` | 1 (`D6`) | 0 | 834 | 1 (`Y111`) | 2 | 0 | Metadato `D6` + representación `Y111` |
| `DetRetri2026072` / `DetValiRetri2026072` | 59 / 63 | 0 | 72 / 69 | 0 | 0 | 0 | Columnas **excluidas por diseño** (`MotivoCadenaFormula`) |
| `Recaudo *` | 48/18/48/16 | 0 | 0 | 0 | 0 | 0 | **Q1 acumulada** por diseño (V3) |
| `ANT EXT-REV` | 34 | 0 | 71 | 0 | 0 | 0 | Hoja **protegida** (la App no la escribe) |
| Resto | **0** | 0 | N | 0 | 0 | 0 | Solo cachés |

**Objetivo H2 cumplido:** `R2` sin la brecha de **53.283.379,41** (suma por bloque `ΔE ≈ 5,9e-8`;
`litVsLit = 0`), `R1 litVsLit = 0`, `R4/SALDOS/banco/AJUSTES/REMUNERACION/GERENTES litVsLit = 0`.
No hay diferencias de contenido nuevas. Las únicas celdas de "contenido" fuera de la lista aceptada
son **ediciones del manual** (`Valida - Control Recaudo!F9=1`, `CONSOLIDADO!Y111` valor,
`VALIDACION_EAAB-CL` columnas D de Q1, `Informe AFaseo!J20="VALIDAR"`): la App no las escribe y no
son derivables de los insumos; se declaran para la propietaria (P5).

### 9.3 Celdas del falso hueco, ahora escritas

| ASE | Celda | Valor | Origen (fuente) |
|---|---|---:|---|
| 2 | `E78` | 1.616.173,63 | `Vlr Servicio` canal 2 |
| 2 | `E79` | 406,42 | `Vlr Intereses` canal 2 (**antes omitido**) |
| 2 | `E80` | 1.616.580,05 | cierre `C=2` |
| 4 | `E313` | 20.152.773,11 | `Vlr Servicio` canal 9 |
| 4 | `E314` | −2.377,86 | `Vlr Intereses` canal 9 |
| 4 | `E315` | 20.150.395,25 | cierre `C=9` (Anticipos) |

### 9.4 Erratas del T0 (corregidas por la medición H2)

1. **§1.5:** el manual v2 **no** conserva fórmulas en `Valida - Control Recaudo!E7/E10/F10`: las
   **convirtió a valores** (68.213.380.915 / 4.284.569.052 / 72.497.949.967) y añadió `F9=1`. La v2
   sí trae las fórmulas `E7=SUM(D7:D9)`, `E10=SUM(D10)`, `F10=E7+E10`; la App las preserva.
2. **§1.2/§1.7:** el `onlyA=1` del `CONSOLIDADO` es **`D111`**: la v1 traía `D111 = D109+N109` y la
   v2 la deja sin fórmula (celda vacía). El pipeline no la usa (la corrida fresca cierra sin ella);
   se certifica su ausencia en el fixture.

### 9.5 `<f>` y proceso

- `diffv` fresh vs `p47-h2c` (`-NoValues`): `fDiff` solo en R2 (352), CONSOLIDADO (61), DetRetri
  (40), DetValiRetri (51), REMUNERACION_* (56-60) y `Valida - Control Recaudo` (3) — **geometría ya
  congelada**; **0** en el resto. La App no crea ni modifica fórmulas.
- Cero `InsertRow`/`DeleteRow`; cero `<f>` nuevas; guard anti-fórmula intacto; CRLF; SHAs v1/v2
  intactos; cero commits.

### 9.6 H1b — cierre de las 7 altas (2026-10-19)

**(a) `Valida - Control Recaudo!E7/E10/F10` (Defecto 1).** En la v2 son fórmulas; el writer escribía
`D7:D10` pero no sellaba sus cachés. H1b añade `SellarCachesControlRecaudo` (patrón Plan 40: el
`<v>` es el valor EVALUADO por read-back, nunca inventado) **después** de escribir `D7:D10`, en los
dos caminos `GenerarWorkbook` (single y multi-ASE). Resultado medido:

| Celda | Valor sellado (julio-2) | Valor manual | Agosto-2 |
|---|---:|---:|---:|
| `E7` | 68.213.380.915 | 68.213.380.915 | — |
| `E10` | 4.284.569.052 | 4.284.569.052 | — |
| `F10` | **72.497.949.967** | 72.497.949.967 | **77.064.563.298** |

`VALIDACION_TOTAL!C14 = 72.497.949.967` y `C15 = 1`. El destino `F10` **no se sobrescribe** (se
conserva el texto `<f>`); el guard anti-fórmula sigue vivo en `D7:D10`.

**(b) `E155` "valor fijo" (Defecto 2) — causa raíz.** **Ningún writer de producción escribe
`E155`**: solo aparece en las listas protegidas. El rojo era del test
`BasePeriodoPreflightTests.S5(2026072)`, que en ese momento corría el pipeline **v2** contra la
**plantilla v1** (`PlantillaUnica()` apuntaba a v1). En la v1 la fila 155 no existe/está vacía; en la
v2 `E155=E101+E123-K101` (fórmula). El guard de la cadena v2 veía la celda v1 como "valor fijo".
Corrección: `PlantillaUnica()` → v2 (base vigente) y `S5` parametriza **Q2→v2 / Q1→v1 histórica**
(Q1 fuera de alcance). Se añadió assert de que el cuadro OPORTUNO conserva sus fórmulas tras la
corrida. **El cuadro OPORTUNO no es escribible** (fórmula o literal-0; §1.3).

**(c) Rebase de tests anclados a v1 (R-FIXTURE, evidencia citada):**
`DivergenciaDeclaradaT0Tests.T0_Oficial_Sha256Intacto` verifica **ambos** SHAs (v1 y v2);
`T0_CorridaVivaQ2_…` usa el escenario sintético + SHA v2; `DivergenciaDeclaradaCierreT4Tests.T4_…`
inyecta el **literal `D7`** (no el caché de `F10`, que es fórmula y el evaluador resuelve);
`ValidaControlRecaudoF10Tests` lee `F10` sellado y añade el test de no-sobrescritura + sellado.

**Suite H1b:** `p48-h1b-suite.trx` = **1059 Total / 1021 Passed / 38 Failed / 0 Skipped**. Set-diff vs
`o43-final-failed.txt` (39): **0 altas / 1 baja heredada** (`ProcesadorPeriodoTests.…VeredictosValidacionesNoVacios`,
causa Plan 46); vs `p47-warnings.trx` (38): **0 altas / 0 bajas**. Build 0 warnings; SHAs v1/v2
intactos; cero commits (`HEAD = c438c8c`).
