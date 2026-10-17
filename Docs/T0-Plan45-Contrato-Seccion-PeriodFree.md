# T0 (Plan 45) — Contrato de sección PERIOD-FREE para filas sin-C + cero documentado

> **Frente:** Plan 45 (H0, criterios S1.1–S1.5). **Naturaleza:** T0 DIAGNÓSTICO — **cero código de producción**.
> **Método:** lectura SOLO con BCL (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`): sin Excel/COM,
> sin OpenXml SDK, sin ExcelDataReader. **El OFICIAL nunca se modifica** (SHA256 antes/después
> `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).
> **Fixture read-only:** `Remuneracion.IntegrationTests/ContratoSeccionPeriodFreeT0Tests.cs` (24 asserts verdes;
> congela por bytes cada hallazgo citado aquí; no toca producción). **Fecha:** 2026-10-16.

---

## 0. Insumos (rutas explícitas)

| Período | Carpeta de insumos | Rol |
|---|---|---|
| Oficial (oráculo/contrato) | `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` | geometría del contrato |
| Julio-2 (2026072, Q2) | `Docs/Prueba Julio-2/Insumos/{ASE}/Recaudoporcomponente_*.xlsx` | evidencia Q2 |
| Agosto (2026082, Q2) | `Docs/Prueba Agosto-2/Insumos/{ASE}/Recaudoporcomponente_*.xlsx` | evidencia Q2 multi-forma |
| Q1 (2026071) | `Docs/Prueba Julio-1/Insumos/{ASE}/Recaudoporcomponente_*.xlsx` | **solo lectura** (citado, no resuelto) |
| Manual legacy ASE5 (medición) | `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx` | hueco de plantilla |

**Prohibido** `Docs/Insumos` raíz. Los cachés `<v>` nunca son oráculo (solo textos-`<f>` y literales).

---

## 1. Encabezados (S1.1 — E1)

- **OFICIAL** (fila 5 ASE1 = 602 ASE5): `F..Z` = `Total · Componente TDF · TTL · TVIAT · Aprovechamiento ·
  CCSA Prest.Aprov. · SERVICIO ESPECIALES · TCS · TLU · TBL · TRT · CCSA Prest. No Aprov. · Ajuste Decena ·
  Ajuste Recargo Aseo · Intereses Cte Financiac. Aseo (T) · Exonera-Recargo Aseo (U) · Exonera-Ajuste Decena
  (V) · Exonera-Recargo Aseo (W) · Exonera-Recargo Aseo (X) · Intereses Cte Financiac. Aseo (Y) · Rec.Intereses (Z)`.
- **FUENTE** (fila 4): `F..U` = los mismos hasta `Intereses Cte Financiac. Aseo (T) · Rec.Intereses (U)`.
- Mapeo por encabezado: `Rec.Intereses` U(fuente) → Z(plantilla); `Intereses Cte` de la fuente → **T**
  (primera ocurrencia); **Y queda intacta** (T/Y no espejo); `Exonera-*` U/V/W/X sin contraparte → intactas.

---

## 2. Fronteras y formas por (período, ASE) (S1.2 — E1/E5)

Las filas-frontera son `B='Mes' C='Total'` y `B='Aplicacion nuevos x reversion' C='Total'`.

**OFICIAL (uniforme, 5/5 ASE):** 5 fronteras en el orden `Mes · Aplic · Mes · Aplic · Mes` → 6 secciones
`[1.ª, Oportuno#1, Aplic#1, Oportuno#2, Aplic#2, Oportuno#3]`.
Filas: ASE1 `16/32/47/59/70`, ASE2 `120/146/171/190/208`, ASE3 `299/315/330/342/353`,
ASE4 `460/486/511/530/548`, ASE5 `613/629/648/660/674`.

**FUENTE (la forma la define el período):**

| Período | ASE | Fronteras (Mes/Aplic en orden) | Forma de secciones |
|---|---|---|---|
| Julio-2 | 1 | Mes,Aplic,Mes,Aplic,Mes | `[1.ª,O#1,A#1,O#2,A#2,O#3]` |
| Julio-2 | 2 | Mes,Aplic,Mes,Aplic,Mes | `[1.ª,O#1,A#1,O#2,A#2,O#3]` |
| Julio-2 | 3 | Mes,Aplic,Mes,Aplic,Mes | `[1.ª,O#1,A#1,O#2,A#2,O#3]` |
| Julio-2 | 4 | Mes,Aplic,Mes,Aplic,Mes | `[1.ª,O#1,A#1,O#2,A#2,O#3]` |
| **Julio-2** | **5** | **Aplic,Mes,Aplic,Mes** | **`[1.ª,A#1,O#1,A#2,O#2]`** (invertida, 5 secciones) |
| Agosto | 1..2 | Mes,Aplic,Mes,Aplic,Mes | `[1.ª,O#1,A#1,O#2,A#2,O#3]` |
| **Agosto** | **3** | **Mes,Mes,Mes** | **`[1.ª,O#1,O#2,O#3]`** (sin Aplicación) |
| Agosto | 4..5 | Mes,Aplic,Mes,Aplic,Mes | `[1.ª,O#1,A#1,O#2,A#2,O#3]` |
| Q1 (read-only) | **1** | **Aplic,Mes,Aplic,Mes** | **`[1.ª,A#1,O#1,A#2,O#2]`** |
| Q1 (read-only) | **2** | **Mes,Mes,Mes** | **`[1.ª,O#1,O#2,O#3]`** |
| Q1 (read-only) | 3..4 | Mes,Aplic,Mes,Aplic,Mes | `[1.ª,O#1,A#1,O#2,A#2,O#3]` |
| Q1 (read-only) | **5** | **Mes,Mes,Mes** | **`[1.ª,O#1,O#2,O#3]`** |

⇒ **Las 3 formas observadas (`[P,O,A,O,A,O]`, `[P,A,O,A,O]`, `[P,O,O,O]`) son SUBSUCESIONES de la
secuencia de tipos del oficial.** No hay una cuarta forma: la fuente elimina o reordena secciones, nunca
inventa tipos.

---

## 3. Regla PERIOD-FREE derivada (S1.4 — E1/E5)

**Emparejamiento monotónico por tipo (greedy):** la k-ésima sección de la FUENTE (en orden) se alinea con
la próxima sección del OFICIAL **del mismo tipo** (§2) aún no consumida. Se computa desde la FORMA
(tipos observados); el contrato no recibe ningún string de período ni de empresa (**D-H**).

| Forma fuente | Mapeo a secciones oficiales |
|---|---|
| `[P,O,A,O,A,O]` | identidad `[1.ª,O#1,A#1,O#2,A#2,O#3]` |
| `[P,A,O,A,O]` (ASE5-julio2, ASE1-Q1) | `[1.ª, A#1, O#2, A#2, O#3]` — **salta O#1 oficial** |
| `[P,O,O,O]` (agosto-ASE3, Q1-ASE2/5) | `[1.ª, O#1, O#2, O#3]` — **salta A#1/A#2 oficiales** |

Las **filas-frontera** viajan con la sección que ABREN: `Mes/Total` de ASE5-julio2 (R33) abre la
`Oportuno#1` fuente → mapeada a `Oportuno#2` oficial → slot **F648**; R54 abre `Oportuno#2` fuente →
`Oportuno#3` oficial → **F674**; el `Oportuno#1` oficial (**F613**) queda sin contraparte (0).

---

## 4. Matriz fuente↔oficial ASE5-julio-2 (S1.3 — caso crítico)

| Sección FUENTE (filas) | Forma fuente | Sección OFICIAL (filas) | Contenido |
|---|---|---|---|
| 1.ª (5–13) | VS,D=H,VS,**VI**,D=O,VS,**VI**,D=T,OCCIDENTE | 1.ª (603–612) | VS,D=E,ENEL,VS,D=H,VS,D=O,VS,D=T,OCCIDENTE |
| A#1 (15–32) | VS/VI + E/ENEL, E/ENERBIT, H, O, T + OCCIDENTE | **A#1 (630–647)** | **idéntica** |
| O#1 (34–39) | `Componente/Subsidio` + D=H, D=O + OCCIDENTE | **O#2 (649–659)** | `Componente/Subsidio` + E/ENEL,H,O,T |
| A#2 (41–53) | `Subsidio` + E/ENEL, E/ENERBIT, H, O, T | **A#2 (661–673)** | **idéntica** |
| O#2 (55–56) | `Subs/Cont`, `Total` | **O#3 (675–676)** | **idéntica** |

- El bloque `Componente/Subsidio` fuente alinea con el **O#2** oficial (649–659), **no** con el O#1
  oficial (615–628, que es un bloque VS/VI con 4 canales). Es la esencia de la regla §3.
- `A#2` y el cierre `Subs/Cont`/`Total` coinciden **exactamente**.

### Colocación corregida por la regla (evidencia del fixture H1)

| Slot | Fuente | Nota |
|---|---|---|
| `F603/F606/F608` | VS de la 1.ª | VS #1..#3 |
| `F607/F609/F611` | D=H / D=O / D=T de la 1.ª | |
| `F612` | R13 OCCIDENTE (1.ª) | 417.519,62 |
| `F604/F610/F616/F620/F623/F626` | — | **0** (la 1.ª fuente no trae estos) |
| `F632` | R17 `D=E/Total` (A#1) | 12.823.196.468,26 (antes mal ubicado en 604) |
| `F648` | R33 `Mes/Total` (frontera O#1→O#2) | 13.084.233.937,44 |
| `F674` | R54 `Mes/Total` (frontera O#2→O#3) | −1.036.320.522,40 |
| `F649/F650/F653/F654/F656/F659` | R34..R39 (`Componente/Subsidio`) | `F659` = R39 OCCIDENTE 24.329,38; **`F628`=0** |
| `F661` | R41 `Subsidio` (A#2) | **≠ `F654`** (duplicado cerrado) |
| `F613` | — | **0** (O#1 oficial sin contraparte) |

---

## 5. Cero documentado (S1.4 — E3/E4)

| Fila | Valor diferido | Destino | Causa |
|---|---|---|---|
| ASE5-julio2 R8 (`O`, `Vlr Intereses`, 1.ª) | **402,40** (+ desglose G..U) | cero documentado | 1.ª oficial sin `Vlr Intereses` |
| ASE5-julio2 R11 (`T`, `Vlr Intereses`, 1.ª) | **7.697,23** (+ desglose) | cero documentado | ídem |

- El **manual legacy** ASE5 sí registraba esas dos filas (`E506=402,4`, `E509=7697.23`) → **hueco de
  PLANTILLA** (no se toca; se reporta al dueño §9).
- **Anti-sobregeneralización (barrido por ASE):** la 1.ª sección del oficial NO trae `Vlr Intereses` en
  NINGÚN ASE (5/5); la fuente solo lo trae en **ASE5-julio2** (ASE1/2/3/4 y agosto no). Los Totales ya
  llevan el dinero: `O` 50.323,46 = 49.921,06 + 402,40; `T` 314.490 = 306.792,77 + 7.697,23.
- **Identidad period-free/company-free:** `(firma: E='Vlr Intereses' con A/B/C/D vacíos, sección: 1.ª)`.
  Cualquier fila sin slot **fuera** de esta lista → `ERR-SIN-SLOT` fail-fast.

---

## 6. Micro-punto F654/F661 (S1.4 — E7)

- **OFICIAL:** `F654` (`D=H`, O#2) = 0 sin `<f>`; `F661` (`Subsidio`, A#2) = 0 sin `<f>`.
- **FUENTE ASE5:** emite **UNA** fila `Subsidio…` (R47, F=−187.855,98) y **UNA** `D=H/Total` (R48,
  F=−187.855,98); el lector (`LeerEspejoR1`) emite cada fila una sola vez.
- **VEREDICTO:** el duplicado **NO nace en el lector** — nace en el **escritor** (cursor físico firma-only
  de las filas sin C). La regla de sección-contrato lo tapa: `Subsidio`/`D=H` van cada uno a su slot de
  sección (no se duplican en 654/661). Re-verificado por el fixture H1.

---

## 7. Formas Q1 citadas SIN resolver (S1.5 — E6; RESTRICCIÓN DURA §0.2)

- **Q1-ASE2:** la 1.ª sección trae el canal **`D=5`** y la empresa **`NUEVO ESQUEMA`**, sin slot en la 1.ª
  oficial (`F110:F119` solo ENEL/OCCIDENTE) → `ERR-SIN-SLOT` (bloqueo Q1 declarado, no se resuelve).
- **Q1-ASE1/5:** formas `[P,A,O,A,O]` / `[P,O,O,O]` (sin `Aplicacion`).
- **Constancia:** Q1 **NO se toca** (ni código ni tests verdes); cualquier cambio exige autorización
  explícita separada. El T0 solo LEYÓ bytes de Q1 (read-only).

**Neutralidad Q1 (decisión del Ingeniero 2026-10-16, opción 1 — implementada en H1):** el escritor
aplica la sección-contrato **solo si el bloque entero es colocable** con ella; si no (Q1-ASE2 trae
`D=5` en la 1.ª sin slot), reproduce el **comportamiento histórico del Plan 41** para ese bloque. Q1
conserva así byte-idéntico su primer `ERR-SIN-SLOT` (`NUEVO ESQUEMA`) y **no se toca ningún test ni
geometría Q1**; set-diff de suite **vacío** contra la línea base de 39. La decisión es por **forma
colocable** (geometría), no por período/empresa (D-H).

---

## 8. Puerta de auditoría (S1.4) — PENDIENTE

La prueba de paridad contra la auditoría manual de la propietaria queda **REDACTADA pero PENDIENTE**
(no llegó). No se declara verde por ausencia (G-AUDITORIA).

---

## 9. Texto listo para reportar el hueco de plantilla al dueño (NO se envía; NO se toca la plantilla)

> En la hoja `Reporte Componentes R1` de la plantilla estandarizada, la **1.ª sección** de cada ASE no trae
> la fila `Vlr Intereses`. Las fuentes `Recaudoporcomponente` del período **sí** traen intereses en esa
> sección (ASE5 julio-2: `O` 402,40 y `T` 7.697,23), y la plantilla anterior de ASE5 también los registraba.
> La app no puede colocarlos porque no existe el slot (no se edita la plantilla). Se requiere confirmar si
> la 1.ª sección debe traer la fila `Vlr Intereses` en todas las ASE.

---

## 10. Veredicto H0 → H1 = **PROCEDE** (la regla absorbe las 3 formas sin ramas)

La regla §3 es geometría pura (recibe la forma, nunca período/empresa): absorbe `[P,O,A,O,A,O]`,
`[P,A,O,A,O]` (inversión ASE5) y `[P,O,O,O]` (sin Aplicación) **sin literales ni ramas**. El mapeo
reproduce el §2.2 del plan (A#1→A#1 oficial; `Componente/Subsidio`→O#2 oficial). **No hay STOP.**

Implementado en H1: `Remuneracion.Core/Models/ContratoSeccionPeriodFreeR1.cs` (alineación) +
`CerosDocumentadosR1.cs` (cero documentado) + consumo en `EscritorValoresSlotR1`. Fixture H1:
`ContratoSeccionPeriodFreeH1Tests` (5 verdes).

---

## 11. Estado fresco (línea base 2026-10-16) — insumo para H2

- **Build:** `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- **Línea base de suite a la entrada (H0):** `39 Failed / 888 Passed / 927 Total` (29 reds pre-existentes +
  8 Q1 declarados + 2 de la clase F682/interior). Fuente: `TestResults/p44-f1-post.trx`.
- **Fixture H0:** 24 tests, todos verdes (no toca producción).
- **Set-diff H2 (post-neutralidad Q1):** suite **39 Failed / 929 Passed / 968 Total**; set-diff **VACÍO**
  contra `o43-final-failed.txt` (mismos 39, 0 nuevos, 0 cerrados); `T1_Ase2Q1` **verde**; oficial SHA256
  intacto; build 0 warnings; sin commits.

---

## 12. ERRATA Plan 46 (2026-10-17) — corrección del contrato de sección y del cero documentado

> **Naturaleza:** errata FECHADA. **NO se reescribe la historia**: los §3/§4/§5/§9 originales quedan como
> registro de lo derivado el 2026-10-16. Lo que sigue es la corrección re-derivada contra BYTES por el
> Plan 46 (`plans/46 - correccion-mapeo-interior-r1-canal-efectivo-y-seccion-por-cierre.md`), cuyo
> fixture read-only es `Remuneracion.IntegrationTests/ContratoSeccionPeriodFreeT0Plan46Tests.cs`
> (26 verdes; matriz citada por ASE × período).

### 12.1 ERRATA §3 (regla PERIOD-FREE) — el tipado es por la frontera que CIERRA

El §3 tipaba la sección por la frontera inmediatamente **ARRIBA** (la que la abre) y trataba la 1.ª
sección como `Primera`. Es incorrecto: **cada sección (fuente y oficial) se tipa por la frontera que la
CIERRA** (la inmediatamente debajo: su fila resumen `Mes`/`Aplicacion…` cuyo importe es su subtotal).
Regla vigente: `ContratoSeccionPeriodFreeR1.SeccionPorCierre` (Core, geometría pura, D-H).

- La 1.ª sección **no** es una categoría fija: si su frontera de cierre es `Aplicacion` se tipa
  `Aplicacion`; si es `Mes`, `Oportuno`.
- Las filas-frontera se tipan por la sección que **ellas cierran** (no por la que abren).
- El emparejamiento `Alinear` sigue siendo monotónico **por tipo** y no cambia.
- La cola sin frontera debajo (`Subs/Cont`, `Total`) se queda con la sección de la frontera de arriba.

### 12.2 ERRATA §4 (matriz fuente↔oficial ASE5-julio-2) — el bloque cae en A#1, no en la 1.ª

La tabla del §4 (y su "colocación corregida") describían el mapeo del tipado por frontera ARRIBA.
Corrección verificada en bytes (fuente `Docs/Prueba Julio-2/Insumos/5-Área Limpia/Recaudoporcomponente_*.xlsx`):

| Sección FUENTE (filas) | Cierre (frontera) | Sección OFICIAL (filas) |
|---|---|---|
| 1.ª (5–13) | `Aplicacion` R14 = 417.519,62 | **Aplic#1 (615–628)** |
| (15–32) | `Mes` R33 = 13.084.233.937,44 | **O#2 (630–647)** |
| (34–39) | `Aplicacion` R40 | **A#2 (649–659)** |
| (41–53) | `Mes` R54 = −1.036.320.522,40 | **O#3 (661–673)** |
| cola (55–56) | — | 675–676 |

- **La 1.ª sección OFICIAL (603–613) y el AFaseo (614) quedan en CERO** (sin contraparte fuente).
- `F612` (OCCIDENTE de la 1.ª oficial) = **0**; el OCCIDENTE de la fuente (R13, 417.519,62) va a **F628**.
- Corrección del §4: donde decía `F603/F606/F608` = VS de la 1.ª, léase **`F619`/`F622`/`F625`**;
  donde decía `F607/F609/F611` = D=H/O/T, léase **`F621`/`F624`/`F627`**; `F612` → **`F628`**;
  `F623` = 402,40 y `F626` = 7.697,23 (antes "0"). `F632` = R17 se mantiene (ya estaba en A#2).

### 12.3 ERRATA §5 (cero documentado) — RETIRADO (WITHDRAW)

Las dos entradas (`Vlr Intereses` de la 1.ª de ASE5-julio2, 402,40 y 7.697,23) **no son un hueco de
plantilla**: con el tipado por cierre caen en su slot real (`F623` / `F626`, dentro de A#1).
Se **retiran** de `CerosDocumentadosR1` (lista de patrones vacía; el MECANISMO de cero documentado se
conserva). El manual legacy ASE5 (`E506/F506`, `E509/F509`) queda como medición de contraste, no como
evidencia de hueco.

### 12.4 ERRATA §9 — WITHDRAW del texto listo para el dueño

El texto del §9 **NO se reporta al dueño**: describe un "hueco de plantilla" que no existe. Queda
retirado y no se envía.

### 12.5 Regla complementaria (Plan 46 §2.1) — CANAL EFECTIVO

Para una fila **sin `D` propio** (y sin `A`/`B`: `Vlr Servicio`, `Vlr Intereses`, `Subsidio…`), su
**canal efectivo** = el `D` de la próxima fila `…/Total` hacia abajo con `D` no vacío y ≠ literal
`Total`. El MISMO cálculo aplica a los slots del oficial. Sin esta regla, las filas sin `D` se asignaban
por cursor de orden dentro de `(firma, sección)` (DEFECTO 1 probado en la corrida 8: `F603`/`F606`/`F608`
recibían VS de H/O/T y `F650`/`F653` los subsidios H/O).

### 12.6 Formas por CIERRE (reemplazan las del §2/§3 para el mapeo)

| Forma por cierre (fuente) | Mapeo a secciones oficiales por cierre |
|---|---|
| `[O#1,A#1,O#2,A#2,O#3]` (julio2 ASE1-4, agosto mayoría, Q1-ASE3/4) | identidad |
| `[A#1,O#1,A#2,O#2]` (julio2 ASE5, Q1-ASE1) | `[A#1(615-628), O#2(630-647), A#2(649-659), O#3(661-673)]` — la 1.ª oficial queda sin contraparte |
| `[O#1,O#2,O#3]` (agosto-ASE3, Q1-ASE2/5) | `[O#1,O#2,O#3]` — saltan A#1/A#2 |

> **Nota de forma `[O#1,O#2,O#3]`:** la fila `Componente` de la fuente vive en una sección cerrada por
> `Mes` (O#3), mientras la `Componente` del oficial vive en la sección cerrada por `Aplicacion` (A#2):
> ese bloque **no es colocable** por el contrato y el escritor reproduce el comportamiento histórico del
> Plan 41 (mecanismo de neutralidad Q1, plan-then-write). `[A#1,O#1,A#2,O#2]` **sí** es colocable.

