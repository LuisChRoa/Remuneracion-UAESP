# T0 (Residual) — Los 24.329,38 de OCCIDENTE ASE5 en el chequeo `O25 = (O24 = N9)`

> **Frente:** residual del Plan 41 / hallazgo del diálogo de reconocimiento. **Naturaleza:** T0 DIAGNÓSTICO —
> **cero código de producción, cero builds de producción**.
> **Método:** lectura SOLO con BCL (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`), sin Excel/COM,
> sin OpenXml SDK. **El OFICIAL nunca se modifica** (SHA256 antes/después
> `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).
> **NO se re-diagnostica** la divergencia declarada OCCIDENTE-ASE5 (`F700` 3 términos vs `F578` 2 términos,
> Δ = 417.519,62): está explicada y ratificada.
> **Fecha:** 2026-10-15.

---

## 0. Insumos (rutas explícitas + SHA256)

| Insumo | Ruta | SHA256 |
|---|---|---|
| Oficial (oráculo de fórmulas) | `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC31…C5214` (intacto) |
| Manual (oráculo recalculado por Excel) | `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx` | `58431010…DCB1` |
| Salida app — **corrida actual (post-Plan 41)** | `Docs/Prueba Julio-2/Resultado/6/Remuneración 202607-2 Total.xlsx` | `B29AC383…1757` |
| Salida app — **corrida pre-fix (defecto viejo)** | `Docs/Prueba Julio-2/Resultado/Resultado5/Remuneración 202607-2 Total.xlsx` | `1CE66248…4F5E` |
| Fuente R1 ASE5 | `Docs/Prueba Julio-2/Insumos/5-Área Limpia/Recaudoporcomponente_…_20268465657830.xlsx` | `963D8C9F…054F` |
| Oráculo externo del administrativo | `Docs/Prueba Julio-2/Insumos/R10_Remuneracion_2026072.xlsx` | `882E7DDA…AE1C` |

> **Nota de trazabilidad:** la tarea nombraba `Resultado5` como “la salida de la app”. `Resultado5` es la
> corrida **pre-fix** (colocación por cursor físico: `F628=229.202.451,62`, `F647=24.329,38`). La corrida que
> produjo el diálogo (hoy, 07:44) es **`Resultado\6`**, y es la que coloca por sección
> (`F628=24.329,38`, `F647=229.202.451,62`, `F673=81.691.821,42`). Todo el análisis usa `Resultado\6`.

---

## 1. El chequeo y los números reportados

`VALIDACION_TOTAL` (oficial, byte-fiel):

```
O25 = =O24=N9
O24 = SUM(O20:O23)      (shared H24:O24, maestro H24 = SUM(H20:H23) → O24 = SUM(O20:O23))
N9  = SUM(N3:N8)
```

| Lado | Origen | Manual (cache recalculado) | App (corrida actual) | Δ (manual − app) |
|---|---|---|---|---|
| `O24` (observado) | `SUM(O20:O23)` = Σ `VALIDACION_{RECIP,ENEL,ENERBIT,OCCIDENTE}!N9` | `73.139.511.002,13` | `73.139.486.672,75` | **24.329,38** |
| `N9` (esperado) | `SUM(N3:N8)` = Σ por-ASE desde `CONSOLIDADO_TOTAL RECAUDO` | `73.139.511.002,13` | `73.139.069.153,13` | **441.849,00** |

- `73.139.511.002,13 − 73.139.486.672,75 = 24.329,38` ✓
- `73.139.511.002,13 − 73.139.069.153,13 = 441.849,00` ✓
- `441.849,00 − 24.329,38 = 417.519,62` ✓ (= Δ declarado)
- App: `O24 − N9 = 73.139.486.672,75 − 73.139.069.153,13 = 417.519,62` ⇒ `O25 = FALSE` **por el Δ declarado**.

**Los dos lados están 24.329,38 por debajo del manual; el lado `N9` está además 417.519,62 por debajo.**
El Δ que el chequeo SÍ ve (417.519,62) es exactamente la divergencia declarada de OCCIDENTE ASE5.
El 24.329,38 es un **desplazamiento común** a ambos lados.

---

## 2. Pregunta 1 — ¿De qué lado está el 24.329,38?

### 2.1 La fila SÍ está colocada (no es hueco de colocación)

Fuente R1 ASE5 (`Sheet1`), filas OCCIDENTE y sus fronteras:

| Fila fuente | C | F | L | Sección (frontera ARRIBA) |
|---|---|---|---|---|
| 13 | OCCIDENTE | `417.519,62` | 0 | 1.ª (antes de `Aplicacion` 14) |
| 14 | Aplicacion nuevos x reversion | `417.519,62` | 0 | (cabecera de Aplic#1) |
| 32 | OCCIDENTE | `229.202.451,62` | `5.904.652,23` | Aplicación #1 (`Aplicacion` 14) |
| 33 | Mes | `13.084.233.937,44` | — | (cabecera de Oportuno#1) |
| 39 | OCCIDENTE | `24.329,38` | 0 | **Oportuno #1 (`Mes` 33)** |
| 40 | Aplicacion nuevos x reversion | `24.329,38` | 0 | (cabecera de Aplic#2) |
| 53 | OCCIDENTE | `81.691.821,42` | 0 | Aplicación #2 (`Aplicacion` 40) |
| 54 | Mes | `−1.036.320.522,4` | — | (cabecera de Oportuno#2) |

Colocación por sección de la corrida actual (`Reporte Componentes R1`, bloque ASE5 = filas 603–676):

| Celda | Rótulo | Valor escrito | Sección |
|---|---|---|---|
| `F612` | `OCCIDENTE` \| Total | `417.519,62` | 1.ª (src13) ✓ |
| `F628` | `OCCIDENTE` \| Total | **`24.329,38`** | Oportuno#1 (src39) ✓ |
| `F647`/`L647` | `OCCIDENTE` \| Total | `229.202.451,62` / `5.904.652,23` | Aplic#1 (src32) ✓ |
| `F659` | `OCCIDENTE` \| Total | `0` | Oportuno#2 (fuente sin fila OCCIDENTE) |
| `F673` | `OCCIDENTE` \| Total | `81.691.821,42` | Aplic#2 (src53) ✓ |
| `F629` | `Aplicacion nuevos x reversion` \| Total | `417.519,62` | (src14) |
| `F660` | `Aplicacion nuevos x reversion` \| Total | `24.329,38` | (src40) |

**Conclusión:** la app coloca la fila `24.329,38` en su slot correcto (`F628`). **No hay hueco de colocación.**

### 2.2 El 24.329,38 es una diferencia ESTRUCTURAL (la fila queda huérfana en el oficial)

Grep exhaustivo de fórmulas en **todas** las hojas del oficial (maestros de fórmulas compartidas incluidos):

```
referencias a F628 / F659 / L628 / L659  →  NINGUNA
```

Es decir: **`F628` (Oportuno#1 de OCCIDENTE ASE5) y `F659` (Oportuno#2) no son consumidos por NINGUNA fórmula
del oficial.** Las únicas anclas del bloque ASE5 son:

| Celda oficial | Texto `<f>` | Consume |
|---|---|---|
| `F680` (TOTAL OPORTUNO → `CONSOLIDADO!D13`) | `F674+F648+F613-L613-L648` | filas `Mes` (613, 648, 674) |
| `F682` (EXTEMPORANEO → `CONSOLIDADO!D51`) | **(literal 0, sin `<f>`)** | nada |
| `F700` (OCCIDENTE OPORTUNO → `REMUNERACION_OCCIDENTE!D13`) | `F647+F673+F612-L612-L647` | `F612` (1.ª), `F647` (Aplic#1), `F673` (Aplic#2) |
| `F702` (OCCIDENTE EXTEMPORANEO → `REMUNERACION_OCCIDENTE!D51`) | **(literal 0, sin `<f>`)** | nada |

Por contraste, en el **mismo oficial** las anclas OCCIDENTE EXTEMPORANEO de ASE1–4 **sí** son fórmulas y
**sí** consumen las filas Oportuno:

| ASE | Ancla EXTEMPORANEO oficial | Filas Oportuno que consume |
|---|---|---|
| 1 | `F97 = F58+F31-L31` | `F31` (Oportuno#1), `F58` (Oportuno#2) |
| 2 | `F276 = F189+F145-L145` | `F145`, `F189` |
| 3 | `F437 = F341+F314-L314` | `F314`, `F341` |
| 4 | `F590 = F529+F485-L485` | `F485`, `F529` |
| **5** | **`F702 = 0` (literal, SIN fórmula)** | **ninguna** |

⇒ El bloque OCCIDENTE ASE5 del oficial es **internamente anómalo**: sus dos filas Oportuno (`F628`, `F659`)
no tienen consumidor, mientras ASE1–4 sí lo tienen. **La app replicó fielmente ese contrato** (su T0 lo
congeló: `GeometriaPlantillaEstandarizadaT0Tests` aserta que `F682` es VALOR sin `<f>`; el rango de slots
del ASE5 es `603–676`, que excluye `F680/F682/F700/F702`).

### 2.3 En el manual la MISMA fila SÍ es consumida (dos veces)

El manual (legacy) tiene DOS familias de anclas y ambas consumen la fila Oportuno#1:

| Ancla manual | Texto `<f>` | Valor | Consumidor |
|---|---|---|---|
| `F580` (OCCIDENTE EXTEMPORANEO) | `F537+F511-L511` | `24.329,38 + 417.519,62 = 441.849,00` | `REMUNERACION_OCCIDENTE_ Directa!D51` |
| `F560` (EXTEMPORANEO) | `F538+F512-L512` | `24.329,38 + 417.519,62 = 441.849,00` | `CONSOLIDADO_TOTAL RECAUDO!D51` |

(`F537`/`F538` = fila Oportuno#1 del manual = fuente fila 39; `F511`/`F512` = 1.ª sección = fuente fila 13.)

---

## 3. Pregunta 2 — Cadena completa hasta el total que entra al chequeo (AMBAS plantillas)

> Correspondencia por **(empresa, ASE, sección)**, NUNCA por dirección (las geometrías difieren:
> bloque ASE5 oficial empieza en 603, legacy en 501).

### 3.1 OFICIAL (corrida actual) — el 24.329,38 NO llega a ningún lado

```
F628 = 24.329,38                      (colocado por la app)
   └─► (sin consumidor: ninguna fórmula referencia F628)
```

**Lado `O24` (espejo de remuneración):**

```
Reporte Componentes R1!F700 = F647+F673+F612-L612-L647      (NO usa F628)
  └─► REMUNERACION_OCCIDENTE_ Directa!D13
        REMUNERACION_OCCIDENTE_ Directa!D108 = D13+D32+D51+D70+D89
          (D51 = 'Reporte Componentes R1'!F702 = 0 literal → NO aporta el 24.329,38)
  └─► VALIDACION_OCCIDENTE!I7 = 'REMUNERACION_OCCIDENTE_ Directa'!D108
  └─► VALIDACION_OCCIDENTE!N7 = SUM(I7:M7)
  └─► VALIDACION_OCCIDENTE!N9 = SUM(N3:N8)
  └─► VALIDACION_TOTAL!O23 = 'VALIDACION_OCCIDENTE'!N9
  └─► VALIDACION_TOTAL!O24 = SUM(O20:O23)          ◄── lado izquierdo del chequeo
```

**Lado `N9` (espejo de recaudo/consolidado):**

```
Reporte Componentes R1!F680 = F674+F648+F613-L613-L648      (NO usa F628)
  └─► CONSOLIDADO_TOTAL RECAUDO!D13
        CONSOLIDADO_TOTAL RECAUDO!D108 = D13+D32+D51+D70+D89
          (D51 = 'Reporte Componentes R1'!F682 = 0 literal → NO aporta el 24.329,38)
  └─► VALIDACION_TOTAL!I7 = 'CONSOLIDADO_TOTAL RECAUDO'!D108
  └─► VALIDACION_TOTAL!N7 = SUM(I7:M7)
  └─► VALIDACION_TOTAL!N9 = SUM(N3:N8)             ◄── lado derecho del chequeo
```

### 3.2 MANUAL (legacy) — el 24.329,38 llega a AMBOS lados

```
F537 = 24.329,38  (Oportuno#1, familia REMUNERACION)
  └─► Reporte Componentes R1!F580 = F537+F511-L511 = 441.849,00
  └─► REMUNERACION_OCCIDENTE_ Directa!D51
  └─► REMUNERACION_OCCIDENTE_ Directa!D108 = D13+D32+D51+D70+D89
  └─► VALIDACION_OCCIDENTE!I7 = 'REMUNERACION_OCCIDENTE_ Directa'!D108
  └─► VALIDACION_OCCIDENTE!N7 = SUM(I7:M7) ─► N9 = SUM(N3:N8) ─► VALIDACION_TOTAL!O23 ─► O24   (lado izquierdo)

F538 = 24.329,38  (Oportuno#1, familia CONSOLIDADO)
  └─► Reporte Componentes R1!F560 = F538+F512-L512 = 441.849,00
  └─► CONSOLIDADO_TOTAL RECAUDO!D51
  └─► CONSOLIDADO_TOTAL RECAUDO!D108 = D13+D32+D51+D70+D89
  └─► VALIDACION_TOTAL!I7 = 'CONSOLIDADO_TOTAL RECAUDO'!D108 ─► N7 = SUM(I7:M7) ─► N9 = SUM(N3:N8)   (lado derecho)
```

**Síntesis:** en el manual la fila entra en los DOS lados ⇒ cancela en `O24=N9` (manual: `O24 = N9 = 73.139.511.002,13`).
En el oficial/app la fila NO entra en NINGÚN lado ⇒ también cancela. **El chequeo es ciego al 24.329,38 en
ambas plantillas.**

### 3.3 Cuantificación de la diferencia de cada lado (app vs manual)

| Celda | Manual | App | Δ |
|---|---|---|---|
| `REMUNERACION_OCCIDENTE!D13` (=`F578` vs `F700`) | `304.989.620,81` | `305.407.140,43` | `+417.519,62` |
| `REMUNERACION_OCCIDENTE!D51` (=`F580` vs `F702`) | `441.849,00` | `0` | `−441.849,00` |
| `REMUNERACION_OCCIDENTE!D108` | `317.721.932,47` | `317.697.603,09` | **`−24.329,38`** |
| `CONSOLIDADO!D13` (=`F558` vs `F680`) | `12.033.011.685,71` | `12.033.011.685,71` | `0` |
| `CONSOLIDADO!D51` (=`F560` vs `F682`) | `441.849,00` | `0` | `−441.849,00` |
| `CONSOLIDADO!D108` | `12.073.344.661,59` | `12.072.902.812,59` | **`−441.849,00`** |

⇒ `O24`: `−24.329,38` (neto de `+417.519,62 −441.849,00`). `N9`: `−441.849,00`. Exactamente lo reportado.

---

## 4. Pregunta 3 — ¿Por qué `O25` es ciego y qué otra validación sí lo vería?

### 4.1 Ceguera de `O25`

`O25` compara **dos lados** (`O24` vs `N9`). Un error que entra o falta en **AMBOS** lados por igual es
**invisible**: solo se detecta el **diferencial** (`O24 − N9`). El 24.329,38 es un desplazamiento de modo común:

- En el manual entra en ambos lados (vía `F580` y `F560`) ⇒ `O24 − N9 = 0`.
- En el oficial/app no entra en ninguno ⇒ `O24 − N9 = 0` (por ese concepto).
- Lo único que `O25` ve es el término `F612` (417.519,62) que el oficial pone SOLO en el lado `O24`
  (vía `F700`) y NO en el lado `N9` (vía `F680`) ⇒ ese es el Δ = 417.519,62 declarado.

**`O25` detecta el diferencial (417.519,62), no el modo común (24.329,38).** Lo mismo aplica a `D34`/`F34`
(comparan `D33`/`F33` contra `CONSOLIDADO!D109`): ambos lados pierden el mismo término.

### 4.2 Validaciones INTERNAS de la plantilla: también ciegas

- `VALIDACION_TOTAL!P3:P9 = INT(O3)=0` con `O3 = H3−N3` (recaudo vs remuneración): el término faltante
  (`CONSOLIDADO!D51 = F682 = 0`) alimenta **a la vez** `N7` (remuneración) y `H7` (vía `D7 = CONSOLIDADO!G129 = D51+N51`),
  así que `O7` no cambia. **Ciego.**
- `VALIDACION_TOTAL!Q28:Q33` (conciliación remuneración): compara agregados internos que pierden el mismo término. **Ciego.**

### 4.3 La validación EXTERNA que SÍ lo ve

`DetRetri` (hoja entregada) re-deriva el total del ASE5 desde las anclas:
`DetRetri!D36 = +D27 + 'Reporte Componentes R1'!F680 + 'Reporte Componentes R1'!F682 + …`
Como `F682 = 0` en el oficial/app, `DetRetri!D36` pierde el 441.849,00. Y el oráculo externo del
administrativo lo tiene:

| Oráculo externo | Celda | Valor |
|---|---|---|
| `R10_Remuneracion_2026072!DetRetri2026072` | `D13` (ASE5) | **`12.073.344.662`** |
| Manual `CONSOLIDADO_TOTAL RECAUDO!D108` | | `12.073.344.661,59` |
| App `CONSOLIDADO_TOTAL RECAUDO!D108` | | `12.072.902.812,59` |

⇒ **El gate `DetRetri-D vs R10 ±0,5` (externo) SÍ vería la diferencia de 441.849,00.** Las validaciones
internas de la plantilla, no.

---

## 5. Pregunta 4 — Impacto en hojas entregadas (cuantificado)

El término faltante **sí** propaga a hojas que el administrativo entrega:

| Hoja entregada | Celda | Manual | App | Δ |
|---|---|---|---|---|
| `REMUNERACION_OCCIDENTE_ Directa` | `D108` (ASE5) | `317.721.932,47` | `317.697.603,09` | **−24.329,38** |
| `REMUNERACION_OCCIDENTE_ Directa` | `D109` (TOTAL) | `4.263.358.205,62` | `4.263.333.876,24` | **−24.329,38** |
| `GERENTES_OCCIDENTE _ Directa` | `D13` (=`REMUNERACION_OCCIDENTE!D108`) | `317.721.932,47` | `317.697.603,09` | **−24.329,38** |
| `CONSOLIDADO_TOTAL RECAUDO` | `D108` (ASE5) | `12.073.344.661,59` | `12.072.902.812,59` | **−441.849,00** |
| `CONSOLIDADO_TOTAL RECAUDO` | `D109` (TOTAL) | `72.441.209.167,71` | `72.440.767.318,71` | **−441.849,00** |
| `GERENTES_TOTAL` | `D13` (=`CONSOLIDADO!D108`) | `12.073.344.661,59` | `12.072.902.812,59` | **−441.849,00** |
| `DetRetri` (entregada) | `D36` (ASE5) | — (usa `F682=0` en app) | pierde `441.849,00` | **−441.849,00** |

**No es “solo un total interno”**: afecta a `REMUNERACION_OCCIDENTE_ Directa`, `GERENTES_OCCIDENTE _ Directa`,
`CONSOLIDADO_TOTAL RECAUDO`, `GERENTES_TOTAL` y `DetRetri` — y las aleja del oráculo externo `R10`.

---

## 6. Veredicto motivado

**(b) Diferencia estructural entre las dos plantillas** — con una precisión de negocio obligatoria.

**Por qué (b):**
1. **No es hueco de colocación.** La app coloca la fila (corrida actual `F628 = 24.329,38`). La colocación
   por sección del Plan 41 funciona.
2. **La ancla/consolidación de esa fila cambió entre plantillas.**
   - Manual: la fila Oportuno#1 (`F537`/`F538`) es consumida por las anclas EXTEMPORANEO (`F580`, `F560`)
     ⇒ entra en AMBOS lados del chequeo.
   - Oficial: la fila (`F628`) **no tiene consumidor** — sus anclas EXTEMPORANEO son **literal 0**
     (`F682`, `F702`), y sus anclas OPORTUNO/TOT_OPT (`F700`, `F680`) no la incluyen.
3. **El propio oficial es internamente anómalo en ASE5**: ASE1–4 traen fórmula EXTEMPORANEO que consume las
   filas Oportuno (`F97`, `F276`, `F437`, `F590`); ASE5 trae `0` literal. La app congeló ese contrato (T0) y
   lo replicó fielmente.

**Precisión de negocio (NO la puede resolver el byte):** si el `EXTEMPORANEO` de OCCIDENTE ASE5 del oficial
(`0` literal) es **deliberado** o una **omisión** de plantilla, es una decisión de negocio de la UAESP —
exactamente la misma naturaleza que la divergencia ya ratificada `F700` (3 términos) vs `F578` (2 términos).
Mientras no se arbitre, la app produce hojas entregadas que difieren del manual **y del R10** en
`24.329,38` / `441.849,00`.

**Matiz crítico sobre “la entrada declarada”:**

> La entrada declarada (Δ = 417.519,62) **NO tapa** el 24.329,38. Tapa el **diferencial** (`O25`/`D34`/`F34`),
> que es justo el término `F612`. El 24.329,38 es **modo común** y ningún chequeo interno de la plantilla lo
> ve; solo lo ve la comparación **externa** contra `R10` (`DetRetri-D vs R10 ±0,5`).

⇒ El veredicto es **(b)** en cuanto a la *naturaleza* (no es defecto de colocación de la app; es la geometría
de consolidación del oficial), **pero** produce una diferencia material no cubierta por la declaración en
hojas entregadas ⇒ **requiere arbitrio de negocio UAESP**, no un “arreglo” unilateral de la app.
**En el encuadre vigente (el manual NO es criterio de validación) el hallazgo se lee como
(B) inconsistencia/omisión de la plantilla oficial** — ver §8.

---

## 7. T0/plan mínimo propuesto (si el Ingeniero decide cerrarlo)

**No se arregla nada en este T0.** Si se aprueba cerrarlo, el plan mínimo sería:

1. **Arbitrio UAESP (bloqueante, 1 pregunta):** ¿el `EXTEMPORANEO` de OCCIDENTE ASE5 del oficial debe ser
   `0` (como hoy) o debe consumir sus filas Oportuno (`F628`/`F659`) como ASE1–4? Dos respuestas:
   - **(i) Deliberado 0** ⇒ no se toca nada; se documenta como divergencia declarada de plantilla y se
     extiende el catálogo de divergencias declaradas (`DivergenciasDeclaradasCatalogo`) para incluir la
     fila `DetRetri!D36`/`CONSOLIDADO!D108` de ASE5, con cita de negocio (no sellar verde).
   - **(ii) Omisión de plantilla** ⇒ Phase 0 **NO puede** escribir `<f>` fuera de allow-list; el cierre
     correcto es que la app **escriba el valor** en las celdas estáticas `F682`/`F702` (son valores, no
     fórmulas) con el `EXTEMP` del dominio (`ExtemporaneoEsperadoPorAse` = `441.849,00`), **sin** insert/delete
     y **sin** tocar el oficial: ampliar el slot-writer del ASE5 para incluir `F682`/`F702` (hoy fuera del
     rango `603–676`) y su gate por (empresa, ASE).
2. **Gate:** `DetRetri-D vs R10 ±0,5` por ASE (ya existe la costura) debe quedar **explícito** para ASE5;
   test que falla si `CONSOLIDADO!D108 ≠ R10!D13`.
3. **Evidencia duradera:** congelar en fixture read-only los bytes citados aquí (`F628` huérfano,
   `F682`/`F702` literales, `F97/F276/F437/F590` fórmulas, `F580`/`F560` manuales, `F628=24.329,38` en la
   corrida actual).

> **Riesgo si se elige (ii) sin arbitrio:** mover el 441.849,00 al lado `N9` re-abre el Δ declarado
> (417.519,62) por otra vía; por eso **(i) vs (ii) es decisión de negocio, no de código**.

---

## 8. Nota de encuadre (reconciliación con la corrección del Ingeniero 2026-10-15)

El Ingeniero fijó (memoria de proyecto `criterio-validacion-app`, 2026-10-15) que **el manual administrativo
es SOLO referencia de desarrollo, NO criterio de validación**: la app debe aplicar las reglas y producir su
salida; el chequeo que falla (`O25`/`D34`/`F34`) **no compara contra el manual**, compara la salida de la app
consigo misma (dos caminos del mismo total). Bajo ese encuadre, este T0 aporta lo siguiente:

- **El hallazgo durable es INTERNO al oficial** y no depende del manual: `Reporte Componentes R1!F628`
  (Oportuno#1 de OCCIDENTE ASE5) y `F659` (Oportuno#2) **no son consumidos por ninguna fórmula del propio
  workbook oficial**, y sus anclas EXTEMPORANEO (`F682`, `F702`) son **literal 0** mientras ASE1–4 traen
  fórmula que sí consume las filas Oportuno. Es decir: **el oficial tiene dos caminos estructuralmente
  distintos** para el mismo total (ancla `F700` que sí usa `F612`, vs ancla `F680` que no) — lo que responde
  la pregunta “¿por qué el propio archivo de la app no cuadra consigo mismo?” con **(B) defecto/omisión de la
  plantilla oficial**, determinable SIN el manual.
- **El manual se usa aquí solo como instrumento de medición del residuo**, no como criterio: sirve para
  cuantificar que la fila Oportuno#1 pesa `24.329,38` y que, donde SÍ tiene ancla (manual), entra en ambos
  lados. La conclusión `(b) estructural` no cambia si el manual se retira de la ecuación.
- **Coherencia con el encuadre:** quien suministró el manual indicó que en OCCIDENTE ASE5 el manual estaba
  mal y la app bien. Este T0 **no contradice** eso: la app coloca la fila correctamente; lo que está en
  discusión es la **ancla de la plantilla oficial**, no la colocación de la app.
- **Investigación hermana en vuelo:** `Docs/T0-ConsistenciaInterna-O25-D34-F34.md` (consistencia interna sin
  el manual). A la fecha de este T0 ese documento **no existe en el repo** ⇒ no se pudo reconciliar su
  contenido; se deja constancia explícita en vez de inferirlo.

---

## 9. Método y reproducibilidad

- Lectura BCL ZIP+XML del oficial, el manual, `Resultado\6`, `Resultado5`, la fuente R1 ASE5 y `R10`, por
  ruta explícita.
- Los textos-`<f>` son el oráculo; los cachés `<v>` solo documentan (nunca demuestran).
- El grep de consumidores se hizo sobre los maestros de fórmula de TODAS las hojas (los followers de
  fórmula compartida no traen texto; un consumo de `F628` habría aparecido en su maestro).
- El oficial se verificó intacto por SHA256 (`3086AC31…C5214`) antes/después.
- Cero `git commit`/`git push`; cero cambios en `Remuneracion.Core`, `Remuneracion.Infrastructure`,
  `Remuneracion.WinForms`, `Remuneracion.Cli`.
