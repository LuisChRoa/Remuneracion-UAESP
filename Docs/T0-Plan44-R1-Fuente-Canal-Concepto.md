# T0 (Plan 44) — Llenado R1 fuente × canal × concepto × componentes: evidencia read-only

> **Frente:** Plan 44 (T0/F0, criterios S1.1–S1.4). **Naturaleza:** T0 DIAGNÓSTICO — **cero código de producción**.
> **Método:** lectura SOLO con BCL (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`): sin Excel/COM,
> sin OpenXml SDK, sin ExcelDataReader. **El OFICIAL nunca se modifica** (SHA256 antes/después
> `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).
> **Fixture read-only:** `Remuneracion.IntegrationTests/R1FuenteCanalConceptoT0Tests.cs` (congela por bytes
> cada hallazgo citado aquí; no toca producción). **Fecha:** 2026-10-16.

---

## 0. Insumos (rutas explícitas + SHA256)

| Insumo | Ruta | SHA256 |
|---|---|---|
| Oficial (oráculo) | `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC31…C5214` (verificado intacto antes/después) |
| Fuente R1 ASE5 | `Docs/Prueba Julio-2/Insumos/5-Área Limpia/Recaudoporcomponente_…_20268465657830.xlsx` | `963D8C9F…054F` |
| Fuente R1 ASE1..4 | `Docs/Prueba Julio-2/Insumos/{1-Promoambiental,2-Lime,3-Ciudad Limpia,4-Bogota Limpia}/Recaudoporcomponente_*.xlsx` | (por ruta) |
| Manual legacy (medición, NO criterio) | `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx` | — |
| Corrida app (defecto medido §0.2 del plan) | `Docs/Prueba Julio-2/Resultado/6/Remuneración 202607-2 Total.xlsx` | — |

**Prohibido** `Docs/Insumos` raíz. Los cachés `<v>` nunca son oráculo (solo los textos-`<f>` y los literales).

---

## 1. Encabezados verificados (S1.1 — E1)

### 1.1 Encabezados F..Z del OFICIAL (fila 5 ASE1 = fila 602 ASE5; bloque `Reporte Componentes R1`)

| Col | Encabezado |
|---|---|
| F | Total |
| G | Componente TDF |
| H | Componente TTL |
| I | Componente TVIAT |
| J | Aprovechamiento |
| K | CCSA Prest.Aprov. |
| L | SERVICIO ESPECIALES |
| M | Componente TCS |
| N | Componente TLU |
| O | Componente TBL |
| P | Componente TRT |
| Q | CCSA Prest. No Aprov. |
| R | Ajuste Decena |
| S | Ajuste Recargo Aseo |
| T | Intereses Cte Financiac. Aseo |
| U | Exonera-Recargo Aseo |
| V | Exonera-Ajuste Decena Aseo |
| W | Exonera-Recargo Aseo |
| X | Exonera-Recargo Aseo |
| Y | Intereses Cte Financiac. Aseo |
| Z | Rec.Intereses |

### 1.2 Encabezados F..U de la FUENTE R1 (fila 4; idéntico en ASE1 y ASE5)

`Total · Componente TDF · Componente TTL · Componente TVIAT · Aprovechamiento · CCSA Prest.Aprov. · SERVICIO ESPECIALES · Componente TCS · Componente TLU · Componente TBL · Componente TRT · CCSA Prest. No Aprov. · Ajuste Decena · Ajuste Recargo Aseo · Intereses Cte Financiac. Aseo · Rec.Intereses`

- La fuente trae **una** columna `Intereses Cte Financiac. Aseo` (col T) y **una** `Rec.Intereses` (col U).
- El oficial trae **dos** columnas con el encabezado `Intereses Cte Financiac. Aseo` (T y Y) — ver §7.
- El oficial trae 4 columnas `Exonera-*` (U/V/W/X) **sin contraparte en la fuente** → se dejan intactas (0 del template).
- Mapeo por encabezado: `Rec.Intereses` U(fuente) → Z(plantilla).

---

## 2. Estructura de secciones — HECHO CENTRAL (S1.2 / S1.3 — E8)

Las filas-frontera de la fuente son `Mes/Total` y `Aplicacion nuevos x reversion/Total`. El **orden y
el número** de fronteras de la fuente NO coincide con el oficial, y **ASE5 es el caso extremo**:

| ASE | Fronteras de la FUENTE en orden (fila) | Nº |
|---|---|---|
| 1 | Mes(11), Aplic(16), Mes(31), Aplic(36), Mes(47) | 5 |
| 2 | Mes(11), Aplic(24), Mes(49), Aplic(59), Mes(77) | 5 |
| 3 | Mes(8), Aplic(14), Mes(29), Aplic(34), Mes(45) | 5 |
| 4 | Mes(11), Aplic(24), Mes(49), Aplic(57), Mes(75) | 5 |
| **5** | **Aplic(14), Mes(33), Aplic(40), Mes(54)** — **empieza con `Aplicacion`, 4 fronteras** | **4** |

| ASE | Fronteras del OFICIAL en orden (fila) | Nº |
|---|---|---|
| 1 | Mes(16), Aplic(32), Mes(47), Aplic(59), Mes(70) | 5 |
| 5 | Mes(613), Aplic(629), Mes(648), Aplic(660), Mes(674) | 5 |

⇒ El oficial tiene SIEMPRE la secuencia `1.ª · Mes · Aplic · Mes · Aplic · Mes(cierre)`. La fuente de
ASE1–4 la reproduce; **la de ASE5 la invierte** (`Aplic` antes del primer `Mes`, y sin el `Mes` de cierre inicial).

**Consecuencia medida:** el emparejamiento `(firma + tipo-de-sección + ordinal)` (§2.1 del plan) alinea
fuente↔oficial en ASE1–4, pero **desalinea en ASE5**:

| Sección | FUENTE ASE5 (filas) | Forma fuente | OFICIAL ASE5 (filas) | Forma oficial |
|---|---|---|---|---|
| 1.ª | 5–13 | VS, D=H, VS, **VI**, D=O, VS, **VI**, D=T, OCCIDENTE | 603–612 | VS, D=E, ENEL, VS, D=H, VS, D=O, VS, D=T, OCCIDENTE (**sin VI**) |
| Aplic#1 (fuente) / Oportuno#1 (oficial) | 15–32 | VS,VI,D=E,ENEL,VS,VI,D=E,ENERBIT,VS,VI,D=H,VS,VI,D=O,VS,VI,D=T,OCCIDENTE | 615–628 | VS,VI,D=E,ENEL,VS,VI,D=H,VS,VI,D=O,VS,VI,D=T,OCCIDENTE (**sin ENERBIT**) |
| Oportuno#1 (fuente) / Oportuno#2 (oficial) | 34–39 | Componente,Subsidio,D=H,Subsidio,D=O,OCCIDENTE | 649–659 | Componente,Subsidio,D=E,ENEL,Subsidio,D=H,Subsidio,D=O,Subsidio,D=T,OCCIDENTE |
| Aplic#2 | 41–53 | Subsidio, D=E, ENEL, … OCCIDENTE | 661–673 | Subsidio, D=E, ENEL, … OCCIDENTE (**idéntica**) |
| Oportuno#2 (fuente) / Oportuno#3 (oficial) | 55–56 | Subs/Cont, Total | 675–676 | Subs/Cont, Total (**idéntica**) |

- Las secciones **`Aplic#2` y el cierre `Subs/Cont`/`Total` de ASE5 coinciden EXACTAMENTE** con el oficial.
- La **`Aplic#1` de la fuente (VS/VI con ENERBIT) coincide EXACTAMENTE** con la `Aplic#1` oficial (630–647) —
  **no** con la `Oportuno#1` oficial (que es un bloque VS/VI sin ENERBIT).
- El bloque `Componente/Subsidio` de la fuente (Oportuno#1, 34–39) coincide estructuralmente con el
  `Oportuno#2` oficial (649–659), no con el `Oportuno#1` oficial.

⇒ **No existe un emparejamiento puro por `(tipo, ordinal)` que lleve cada fila fuente de ASE5 a su slot
de forma correcta y sin huecos.** La forma del bloque (no la etiqueta de frontera) es la que manda, y la
fuente trae una sección `Mes` de menos que el oficial.

---

## 3. Colocación actual medida (desvío (b) del §0.2) — reproducida 11/11

Con la colocación vigente (veto de sección **solo** en filas-dato de empresa; filas sin C por cursor
físico **firma-only**), las 11 celdas del §0.2(b) quedan así (todas reproducidas desde las filas-fuente):

| Celda slot | Valor | Fila fuente de origen | Sección fuente | Sección del slot |
|---|---|---|---|---|
| `F604` | 12.823.196.468,26 | R17 `D=E/Total` | Aplic#1 | 1.ª |
| `F610` | 12.815.008.425,00 | R15 `Vlr Servicio` | Aplic#1 | 1.ª |
| `F616` | 402,40 | R8 `Vlr Intereses` | 1.ª | Oportuno#1 |
| `F620` | 7.697,23 | R11 `Vlr Intereses` | 1.ª | Oportuno#1 |
| `F623` | 8.188.043,26 | R16 `Vlr Intereses` | Aplic#1 | Oportuno#1 |
| `F626` | 90.800,20 | R20 `Vlr Intereses` | Aplic#1 | Oportuno#1 |
| `F631` | 203.478,11 | R24 `Vlr Intereses` | Aplic#1 | Aplic#1 |
| `F635` | 1.803.780,95 | R27 `Vlr Intereses` | Aplic#1 | Aplic#1 |
| `F639` | 1.037.886,13 | R30 `Vlr Intereses` | Aplic#1 | Aplic#1 |
| `F613` | 13.084.233.937,44 | R33 `Mes/Total` | Aplic#1 | Mes#1 |
| `F648` | −1.036.320.522,40 | R54 `Mes/Total` | Oportuno#2 | Mes#2 |

**Las filas con empresa en C quedan CORRECTAS** (colocación por empresa + sección): `F612`=417.519,62 (R13),
`F628`=24.329,38 (R39), `F647`=229.202.451,62/`L647`=5.904.652,23 (R32), `F673`=81.691.821,42 (R53).

**VEREDICTO (b):** el defecto vive **enteramente** en el cursor físico de las filas sin C. La regla del
§2.1 **NO lo corrige** (el veto de sección no aplica a las filas sin C, así que no cambia nada); y
extenderlo a las filas sin C (VS/VI/D-code/Subsidio) **rompe ASE5** (§2): la `Vlr Intereses` de la 1.ª
sección de ASE5 no tiene slot en la 1.ª sección oficial, y el bloque `Componente/Subsidio` cae en una
sección oficial de forma distinta.

---

## 4. Hueco de PLANTILLA: 1.ª sección sin `Vlr Intereses` (S1.1/S1.3 — E2/E3/E4)

| ASE | 1.ª sección OFICIAL | ¿trae `Vlr Intereses`? | 1.ª sección FUENTE | ¿trae? |
|---|---|---|---|---|
| 1 | 6–15 | **NO** (4 VS + D=E/H/O/T + ENEL/OCCIDENTE) | 5–10 | NO |
| 2 | 110–119 | **NO** | — | (por confirmar por ASE en el barrido) |
| 3 | 289–298 | **NO** | — | — |
| 4 | 450–459 | **NO** | — | — |
| 5 | 603–612 | **NO** | 5–13 | **SÍ**: R8 (`O`) F=402,40 y R11 (`T`) F=7.697,23, con desglose G..U |

El manual legacy ASE5 (bloque 501) **sí** registraba esas dos filas: `E506=Vlr Intereses`/`F506=402,4` y
`E509=Vlr Intereses`/`F509=7697.23`.

**Hueco de PLANTILLA (no de la app):** la fuente de ASE5 trae valores en la 1.ª sección que el oficial
no tiene dónde recibir. **NO se toca la plantilla** (SHA256 intacto) — se **reporta al dueño** (texto listo
en §9). Nota anti-sobregeneralización: en ASE1 la fuente tampoco trae intereses en la 1.ª (coincide con la
plantilla); el hueco es **por ASE** (al menos ASE5 lo tiene).

**CONTRADICCIÓN INTERNA DEL PLAN (bloqueante):** §2.1 manda `ERR-SIN-SLOT` para toda fila fuente sin slot;
pero §0.2(c) y el criterio de salida de F1 (§1.2) exigen "fila sin contraparte queda en cero, sin
`ERR-SIN-SLOT` espurio" y "no se parchea la plantilla". Ambas reglas no pueden regir a la vez para las
filas `Vlr Intereses` de la 1.ª de ASE5.

---

## 5. Micro-punto `F654`/`F661` — aislamiento lector vs escritor (S1.3 — E6)

- **Oficial:** `F654` (Oportuno#2, `D=H`) = 0 sin `<f>`; `F661` (Aplic#2, `E=Subsidio(-)/Contribucion(+)`) = 0 sin `<f>`.
- **Fuente ASE5:** emite **UNA** fila `Subsidio…` (R47, `F=−187.855,98`) y **UNA** fila `D=H/Total` (R48, `F=−187.855,98`).
  El lector (`LeerEspejoR1`) emite cada fila fuente **una sola vez**.
- Con el cursor físico firma-only: las 4 filas `D=H/Total` de la fuente caen en los 4 primeros slots
  `D=H` del oficial (607, 621, 640, **654**) y las 7 filas `Subsidio` en los 7 primeros slots `Subsidio`
  (650, 653, 655, 657, **661**, 664, 667). De ahí `F654 = F661 = −187.855,98`.

**VEREDICTO:** el duplicado **NO nace en el lector** (una emisión por fila) — nace en el **escritor** por el
cursor físico de las filas sin C. La regla sin-C de F0 (I o F) **sí puede** y **debe** taparlo; la
condición del §2.2 ("si el duplicado naciera en el lector, la regla no lo tapa") **no aplica**.

---

## 6. Anclas OCCIDENTE ASE5 / `F628` huérfana (S1.1 — E5; contexto F2, no se re-decide)

Textos-`<f>` byte-fiel del oficial:

| Celda | Texto `<f>` | Naturaleza |
|---|---|---|
| `F700` (OCCIDENTE OPORTUNO) | `F647+F673+F612-L612-L647` | 3 términos (incluye la 1.ª sección) |
| `F680` (TOTAL OPORTUNO → `CONSOLIDADO!D13`) | `F674+F648+F613-L613-L648` | 3 términos `Mes` |
| `F682` (EXTEMPORANEO) | — | **literal 0** |
| `F702` (OCCIDENTE EXTEMPORANEO) | — | **literal 0** |

Grep exhaustivo de consumidores (todas las hojas, maestros shared incluidos): **ninguna** fórmula
referencia `F628`/`F659`/`L628`/`L659`. Es decir, las filas Oportuno de OCCIDENTE ASE5 **no tienen
consumidor**, mientras ASE1–4 sí lo tienen (`F97`, `F276`, `F437`, `F590`). Coherente con
`Docs/T0-Residual-24329-OccidenteAse5.md` (residual 24.329,38) y `Docs/T0-Plan41-MiniT0-OccidenteAse5-…md` §1.

**No se re-decide aquí** la vía A/B: eso es F2, y exige números frescos **tras F1**.

---

## 7. Veredicto T/Y (S1.3 — E7)

- El oficial declara **dos** columnas con el encabezado idéntico `Intereses Cte Financiac. Aseo` (T y Y).
- El **ancla de empresa del propio oficial** consume ambas por separado, p. ej. (cita literal en
  `REMUNERACION_OCCIDENTE_ Directa`):
  `'Reporte Componentes R1'!T95+'Reporte Componentes R1'!W95+'Reporte Componentes R1'!X95+'Reporte Componentes R1'!Y95+'Reporte Componentes R1'!Z95+'Reporte Componentes R1'!S95`.
- Si Y fuera **espejo** de T, el ancla sumaría el interés **dos veces** (el template sería incoherente
  consigo mismo). Como el ancla las trata como **términos distintos**, Y **NO es espejo**.

**VEREDICTO T/Y:** **no espejo**. El escritor multicolumna debe escribir el `Intereses Cte` de la fuente
**solo en T** (primera ocurrencia del encabezado) y dejar **Y intacta** (0 del template). Además, el mapeo
"por encabezado" del §2.3 es **no inyectivo** en la plantilla (T y Y comparten texto): el T0 fija que la
resolución por encabezado se hace por **(texto, primera ocurrencia)**, y que Y queda fuera del contrato.

---

## 8. Puerta de auditoría (S1.4) — PENDIENTE (no verde por ausencia)

- Geometría y ejemplos congelados en el fixture (§0–§7).
- La **prueba de paridad contra la auditoría manual de la propietaria** queda **REDACTADA pero PENDIENTE**:
  la auditoría **aún no fue entregada**. **No se declara verde por ausencia** (G-AUDITORIA del §S0).
- Cuando llegue, se ejecuta como oráculo (nuevo T0 de reconciliación si discrepa de lo congelado aquí).

---

## 9. Texto listo para reportar el hueco de plantilla al dueño (NO se envía; NO se toca la plantilla)

> En la hoja `Reporte Componentes R1` de la plantilla estandarizada, la **1.ª sección** de cada ASE no trae
> la fila `Vlr Intereses`. Las fuentes `Recaudoporcomponente` del período **sí** traen intereses en esa
> sección (p. ej. ASE5 julio-2: `O` 402,40 y `T` 7.697,23), y la plantilla **anterior** de ASE5 también los
> registraba. La app **no puede** colocarlos porque no existe el slot (y no se edita la plantilla). Se
> requiere confirmar si la 1.ª sección debe traer la fila `Vlr Intereses` en todas las ASE.

---

## 10. VEREDICTO F0 → F1 = **STOP** (con causa citada)

No se abre T1 (F1) con la especificación vigente del plan, por **dos causas byte-probadas**:

1. **(Placement rule indefinida — §2/§3).** La clave `(C-por-sinonimia, D, E, sección-por-frontera)` del
   §2.1 resuelve las **filas-dato de empresa** de ASE5 correctamente, pero **no** las filas sin C: el
   emparejamiento por `(tipo, ordinal)` **no existe** para ASE5 (su fuente invierte el orden de fronteras y
   trae una sección `Mes` de menos). Aplicar la sección a las filas sin C deja filas **sin slot**
   (p. ej. la `Vlr Intereses` de la 1.ª) → `ERR-SIN-SLOT` sobre los datos de julio-2 que F1 debe cerrar.
   Ni la Opción I ni la Opción F del §2.2 están definidas sobre bytes reales; y **§2.1, tal como está
   redactada, no corrige el desvío (b)**.
2. **(Contradicción de la fila sin slot — §4).** §2.1 exige `ERR-SIN-SLOT`; §0.2(c)/§1.2 exigen "cero sin
   `ERR-SIN-SLOT` espurio" y "no parchear la plantilla". Mutuamente excluyentes para el hueco de la 1.ª
   (`Vlr Intereses`).

**Desbloqueado por F0 (decisiones SÍ tomadas):**
- **T/Y:** no espejo; writer solo T; Y intacta; mapeo por `(texto, primera ocurrencia)` (§7).
- **Micro-punto F654/F661:** artefacto de **escritura** (cursor físico firma-only), no del lector; la regla
  sin-C debe taparlo (§5).
- **Mapeo U(fuente)→Z(plantilla)** confirmado por encabezado; `Exonera-*` U/V/W/X del oficial intactas (§1).
- **Anclas/contexto F2** congelados (§6); vía A/B se decide **solo** en F2 con números frescos.

**Bloqueado:** F1 (implementación multicolumna + regla sin-C), F2 (re-verificación tras F1) y F3 (corridas
frescas tras F1). La **auditoría** sigue PENDIENTE (§8).

---

## 11. Estado fresco pre-F1 (línea base 2026-10-16) — insumo para F2

- **Build:** `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- **Suite completa** (`dotnet test … --logger "trx"`, 2026-10-16): **39 Failed / 888 Passed / 0 Skipped / 927 Total**.
  - **Set-diff de fallos = VACÍO** contra la línea base de entrada (`TestResults/o43-final-failed.txt`, 39 nombres): **mismos 39**.
  - Nota: el plan cita `39 Failed / 873 Passed / 912 Total` (Plan 42 T5); el repo evolucionó a 927 tests (Plan 43). El criterio vinculante (mismos 39 fallos) se cumple. El fixture F0 añade **12 tests** (todos verdes).
- **Autocontroles W-1 frescos (Q2 2026072, del trx):**

| Celda | Estado | Detalle |
|---|---|---|
| `VALIDACION_TOTAL!C15` | **TRUE** | recaudo cierra |
| `VALIDACION_TOTAL!D25` | **TRUE** | recaudo cierra |
| `VALIDACION_TOTAL!O25` | **FALSE** | `ERR-VALIDACION` en ASE1..5 — divergencia **declarada** `OCCIDENTE-ASE5`, Δ = **417.519,62** |
| `VALIDACION_TOTAL!D34` | **FALSE** | ídem (Δ = 417.519,62) |
| `VALIDACION_TOTAL!F34` | **FALSE** | ídem (Δ = 417.519,62) |

- **Gates R1 frescos:**
  - `ValidadorTotalesR1GateTests.S1_Julio2026072_GateVerde`: ROJO — `ASE 5 R1!F682 [EXTEMP]: esperado 441849.00, evaluado 0` (residual §6; `F682` está **fuera** del rango de slots 603–676).
  - `ValidadorTotalesR1GateTests.S1_Agosto2026082_GateVerde`: ROJO — `ASE 5 R1!F682 [EXTEMP]: esperado 1249810.00, evaluado 0` (misma clase).
  - `ValidadorInteriorR1GateTests.S1_{Julio2026072,Agosto2026082}`: ROJO — la **forma interior** de R1 no cierra (celdas interiores sin `<f>`); frente de recomposición interior (ajeno a este plan).
- **Q1 (2026071):** `RegresionEvaluaSellaPlan40Tests.T4` ROJO por `ERR-SIN-SLOT: ASE 2` (bloqueo Q1 declarado, no se toca).

> Esta tabla es el **punto de partida** de F2. F2 exige re-correr todo **tras F1**; F1 está STOP (§10), así que estos
> números **no** constituyen el cierre de F2.

---

## 12. Propuesta para el Ingeniero (ratificar antes de reabrir F1)

La evidencia apunta a que la clave correcta **no** es `(firma + tipo/ordinal de frontera)` sino una
alineación **por forma de sección** (el oficial tiene una `Mes` de cierre extra y ASE5 invierte el orden).
Opciones concretas a arbitrar:

- **(P1) Alineación por forma/ordinal-de-sección (sin tipo):** emparejar la k-ésima sección **no vacía**
  fuente↔oficial por la **firma del representante**, tratando las secciones `Mes` como consecutivas, y
  `Vlr Intereses` sin slot como **cero** (documentado), no `ERR-SIN-SLOT`.
- **(P2) Contrato de sección Q2-aware por ASE:** declarar en el contrato T0, por ASE, el mapeo explícito
  sección-fuente↔sección-oficial (5 ASE × ~6 secciones), con cita byte. Cero ramas por período; el contrato
  es geometría, no negocio.
- **(P3) Aceptar el algoritmo vigente (cursor físico firma-only) como contrato** y limitar F1 al llenado
  multicolumna `F..Z` **sin** corregir (b) — sólo válido si el Ingeniero declara (b) "no-defecto".

Se recomienda **P1 o P2** con ratificación del Ingeniero (Clarification Gate §7 del plan, punto 2).

---

## 13. Método y reproducibilidad

- Lectura BCL ZIP+XML del oficial, las 5 fuentes R1, el manual legacy, por ruta explícita.
- Textos-`<f>` y literales son el oráculo; los cachés `<v>` solo documentan (nunca demuestran).
- El grep de consumidores se hizo sobre los `<f>` de TODAS las hojas (maestros shared incluidos).
- Oficial verificado intacto por SHA256 antes/después (`3086AC31…C5214`).
- Cero `git commit`/`git push`; cero cambios en `Remuneracion.Core`, `Remuneracion.Infrastructure`,
  `Remuneracion.WinForms`, `Remuneracion.Cli`.
