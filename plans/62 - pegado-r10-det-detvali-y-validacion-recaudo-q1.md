---
titulo: Plan 62 - Pegado R10 DetRetri/DetValiRetri y validacion del recaudo por quincena en CONSOLIDADO 142-147
estado: APROBADO 2026-10-28 — el Ingeniero ratificó las 8 decisiones del Clarification Gate §6; H0 abierto
fecha: 2026-10-28
numero: 62
planes_ocupados: 01..61 ocupados; este es el 62
padre: brief del Ingeniero 2026-10-28 (novedades septiembre-1: N1 pegado R10 + N2 CONSOLIDADO 142-147 por quincena)
madre_doctrina: .opencode/project-context.md (plantilla autoridad; values-only; guard anti-fórmula; tolerancia ±0,5; prohibido `<v>` stale como oráculo; 0 warnings; sin valores quemados por período/ASE/canal; nunca sobrescribir fórmulas) + plans/59 (asiento DetRetri H4, geometría derivada, candado anti-quemado) + plans/61 (frente R10, ClasificadorNovedadesEntrega, entrega con novedades) + plans/40 (familia E1-4, Phase 0 por dominio) + plans/31 (sello de proceso D6/D7) + plans/29 (comparador, exclusiones versionadas) + plans/30 (NombresHojaPeriodo)
caso_rojo: corrida septiembre-1 (2026091/Q1) — `ValidadorBasico.ValidarDetRetriContraR10` + frente R10 del Plan 61 (`ClasificadorNovedadesEntrega`) reportó `-1` en el TOTAL (Σ per-ASE calculado vs total declarado R10; cada per-ASE cuadra ±0,5). Causa: divergencia de redondeo de composición. El proceso manual del administrativo es copiar el bloque R10 a DetRetri/DetValiRetri; la salida de la app debe ser indistinguible de ese manual.
gate_entrada:
  plantilla_oficial: Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx — v1 citada; el pipeline consume `Insumos.PlantillaOficialV4` → Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx (VIGENTE; verificado H4: geometría E142:E146/L145/M142 idéntica v1-v4). Huella distinta ⇒ STOP
  r10_oraculo: Docs/Prueba Septiembre 1/Insumos/R10_Remuneracion_2026091.xlsx — 2 hojas (`DetRetri2026091`, `DetValiRetri2026091`)
  base_paridad: Docs/Prueba Septiembre 1/Resultados/Administrativo/Remuneracion 202609-1 Total.xlsx (manual del administrativo, oráculo de paridad del pegado)
  suite_base: la vigente al abrir H0 (trx de entrada citado en H0; 0 Skip; set-diff: rojo nuevo fuera de lista = STOP)
  sin_commits: NUNCA git commit / git push; NUNCA git checkout / git reset
  excel_cerrado: Excel debe estar cerrado durante las corridas
---

# Plan 62 — Pegado R10 DetRetri/DetValiRetri y validación del recaudo por quincena en CONSOLIDADO 142–147

> SDD completo PROPOSE → DESIGN → SPEC → TASKS → GATES. Fuente única de verdad para la implementación una vez aprobado ESTE documento. No se implementa sin aprobación del Ingeniero (punto de bloqueo: sin su OK no se abre H0; estado inicial PROPUESTO).
> Lectura obligatoria previa: `plans/59` (asiento DetRetri H4, §3.4 slots literales, geometría derivada) + `plans/61` (frente R10, entrega con novedades, malla mínima por período) + `plans/40` (familia E1-4, Phase 0 por dominio, F-NEW-1) + `plans/31` (sello de proceso D6/D7) + `.opencode/project-context.md` (doctrina madre).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/Cli/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español técnico; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan; WinForms/CLI solo heredan el estado y el código de salida del Plan 61).

---

## §0. Breadcrumb de patrones y citas a planes previos (qué se reutiliza, qué cambia)

| # | Patrón / doctrina previa | Uso en este plan |
|---|---|---|
| P-59-H4 | Asiento DetRetri (Plan 59 H4: `EscribirCeldasDetRetri` D9:D13/D14 + `EscribirDesgloseDetRetri` J/L/Cierre en `OpenXmlPlantillaWriter.cs:337/347/1113-1180/1748-1835` + `VerificarAsientoDetRetri` + "23 celdas" de `:928`) | **RETIRADO** por el pegado (puntos 8/9 del brief): las celdas que escribía caen dentro de D9:O13 y J9:J14 ⇒ sobre-escritas por el pegado. H2 traza consumidores antes de retirar (T2.5) |
| P-59-§3.4 | Slots de montos DetRetri D9:O14 son literales `<v>0</v>` (destinos de escritura); filas 3–19 sin `<f>`; fórmulas D23:D28/D32:D36 + auto-chequeos L32:O36 bajo la fila 19 | Base del pegado: el rango A3:O19 es values-only por construcción ⇒ el guard anti-fórmula sigue como red (tercera forma = `ERR-PLANTILLA`) |
| P-61-R10 | `ValidarDetRetriContraR10` (AMBAS quincenas, `ProcesadorPeriodo.cs:443`) + `ClasificadorNovedadesEntrega.DesdeR10` (Plan 61, entrega con novedades) + `ExcelDataReaderDetRetriR10Reader` (conversor monetario texto `$…`/negativos) | El conversor se REUTILIZA (extraído a fuente única, no duplicado). El gate total contra R10 queda **OBSOLETO para la hoja** (tautología) y se sustituye por verificación de integridad del pegado (H3) |
| P-40-E1-4 | `PreparadorBasePipeline.cs:112-130` (`CeldasRecaudoValidacion` C3:C8 → filas 21:26 + `ColumnaRecaudoPorQuincena` {1→D, 2→F} + `ColumnaRecaudoDe(periodo)`) + reescritura `AplicarVarianteRecaudoValidacion` (~línea 543) + familia `Verificar*` (~línea 862+) | **EXTENDIDA** con la misma mecánica a `E142:E146 → filas 21:25` (H4). El ejemplo D24 del negocio describe su archivo manual de Q1, NO la geometría de la plantilla (filas 21:25 por doctrina Plan 40) |
| P-31-D-F | Sello de proceso `DetRetri*/DetValiRetri*!D6/D7` (texto dd/MM/yyyy + hh:mm AM/PM desde el R10, `ProcesadorPeriodo.cs:355-379` + `SellarProcesoEnDetRetri` en `OpenXmlPlantillaWriter.cs:2570-2640`) | **RETIRADO** (decisión D-RET-SELL0 §2): el pegado A3:O19/A3:O14 ya trae D6/D7 del R10 por construcción. Menos escrituras activas, cero estados intermedios |
| P-29-D-E | `ComparadorSalidaVsManual` (BCL puro, f-vs-f + literales ±0.5, prohibido `<v>` stale) + exclusiones versionadas (incluida `DetValiRetri!J9` = SALE) | `DetValiRetri!J9` (`AJUSTE A LA DECENA`) **vuelve al comparador**: entra al pegado ⇒ la exclusión SALE del catálogo (H5) |
| P-30-D-A | `NombresHojaPeriodo.DetRetri/DetValiRetri(codigoCompleto)` — cero literales de hoja en runtime | El pegado resuelve AMBAS hojas por el helper (origen y destino compuestos del dominio; sin detección por enumeración) |
| P-28 | `fullCalcOnLoad` + sin `calcChain` (`SaneadorCadenaCalcula.Sanear` en los 3 caminos de guardado) | Intacto: el pegado solo toca `<v>`/texto, nunca `<f>`; la salida sigue abriendo sin reparación |
| G3-D1 | "R10 = oráculo NUNCA se escribe" (project-context: `R10 … Oráculo de VALIDACIÓN … NUNCA se escribe`) | **REEMPLAZADA** por N1 (ver §Doctrina D-G3): el R10 pasa a ser FUENTE de DetRetri/DetValiRetri en TODOS los períodos y AMBAS quincenas |

---

## §1. PROPOSE — Problema / Objetivo

### 1.1 Intent

Que la salida de la app sea **indistinguible del archivo manual del administrativo** en los dos frentes que septiembre-1 dejó en rojo:

1. **N1 — DetRetri/DetValiRetri se copian del R10.** La app pega el bloque completo del insumo `R10_Remuneracion_{AAAAMMQ}.xlsx` (hojas `DetRetri{AAAAMMQ}` filas 3–19 cols A–O y `DetValiRetri{AAAAMMQ}` filas 3–14 cols A–O) en las hojas homónimas de la plantilla de procesamiento, en TODOS los períodos y AMBAS quincenas. Con el pegado, el `D14` entregado ≡ R10 por construcción ⇒ la novedad de total `-1` DESAPARECE como reporte.
2. **N2 — CONSOLIDADO 142–147 varía por quincena.** `CONSOLIDADO_TOTAL RECAUDO!E142:E146` (bloque "VALIDACION DEL RECAUDO VERSUS REMUNERACION") se reescribe por `NumeroQuincena` igual que la familia E1-4: Q1 → columna D, Q2 → columna F de las 5 hojas `Recaudo *` (filas 21:25 de la plantilla).

### 1.2 In Scope (H0 + H1 + H2 + H3 + H4 + H5 + Gate, prohibido ampliar sin nuevo OK del Ingeniero)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **H0/T0** | Verificación read-only: re-medición de los 12 hechos del brief + asserts blindados | Fixture verde + doc de evidencia; puerta H0→H1 |
| **H1** | Lector de bloque R10 (Core interface + Infrastructure ExcelDataReader, reutilizando el parser monetario) | CA-C1 en 2026091 + goldens sin regresión |
| **H2** | Writer de pegado OpenXML (A3:O19 / A3:O14, values-only, estilos de destino preservados, guard anti-fórmula) + retiro de escrituras solapadas (P-59-H4, sello Plan 31, rótulos `EscribirDesgloseTrazableDetRetri`) con traza de consumidores | CA-C2 + 0 escritura muerta |
| **H3** | Reformulación de gates: per-ASE conservado (o equivalente DetValiRetri D16:D20 ≈ 0 ±0.5) + gate total sustituido por verificación de integridad del pegado (read-back ≡ R10, fail-fast con hoja/celda/valor) | CA-C3: septiembre-1 sin novedad R10 `-1`; `D21` residual vive donde ya lo tiene el manual (divergencia documentada, excluida del gate — D6) |
| **H4** | Extensión E1-4 al bloque 142–147 (tabla `CeldasRecaudoValidacionConsolidado142` + reescritura por quincena + verificación base) | CA-C4: Q1 → D{21:25}, Q2 → F{21:25}, `fullCalcOnLoad` alcanzable |
| **H5** | Reclasificación de tests/comparador (exclusión J9 OUT, asserts "fórmulas Q1 filas 9–14 intactas" → doctrina nueva) + actualización de `.opencode/project-context.md` + manual de usuario | CA-C5 + suite verde con espectro de reclasificación explícito |
| **Gate** | Corridas 2026091 (caso), 2026071/2026072 + goldens Julio Q1/Q2 (no-regresión con solo desplazamientos DECLARADOS) + suite completa con set-diff | CA-C6 |

### 1.3 Non-goals / out (vinculante)

- Cualquier otra hoja del workbook; cambios de negocio en fórmulas fuera de `E142:E146`.
- La deuda `R-SLOT-Q1-ENERBIT` (Planes 41/46), F-NEW-1 hardcode M (requiere decisión), RecaudosReversados.
- Tocar la plantilla oficial (huella intacta; cero `<f>` nuevo, cero `InsertRow`/`DeleteRow`).
- Commits, docs de usuario fuera de lo citado. NUNCA `git commit` / `git push`. NUNCA `git checkout` / `git reset`.

### 1.4 Riesgos con mitigación (resumen; detalle en §7)

| ID | Riesgo | Mitigación |
|---|---|---|
| R-PEGADO-FORMULA | El destino trae `<f>` en A3:O19/A3:O14 (plantilla futura) y el pegado la pisa en silencio | D-H2: guard anti-fórmula pre-escritura por celda (tercera forma = `ERR-PLANTILLA` nombrando hoja+celda) |
| R-DOBLE-ESCRITURA | Escrituras solapadas (H4 Plan 59, sello Plan 31, rótulos) siguen corriendo "por debajo" y dejan estados intermedios confusos | D-H2: retiro con traza de consumidores (T2.5); si un consumidor necesita el valor, el valor sigue en el dominio (Inputs), no en la hoja |
| R-TAUTOLOGIA | Conservar el gate total calculado-vs-R10 tras el pegado = verde que no prueba nada | D-H3: el gate total se SUSTITUYE por verificación de integridad read-back; el per-ASE se conserva (o su equivalente D16:D20) |
| R-QUINCENA-FIJA | Congelar D o F en `E142:E146` en vez de gobernar por dominio | D-H4: tabla por `NumeroQuincena` reutilizando `ColumnaRecaudoDe(periodo)`; tercera quincena = `ERR-PLANTILLA` |
| R-LETRA-FISICA | Resolver Especiales/columnas del R10 por letra en vez de por encabezado | D-H1: localización por etiquetas (header fila 8 `ASE`/`RECAUDO TOTAL`); la letra física NUNCA se nombra en producción |
| R-ORACULO-STALE | Comparar con cachés `<v>` del manual o de la salida | G-WORKBOOK (Plan 28/29): texto-`<f>` + literales; `<v>` de fórmula nunca es oráculo |
| R-REGRESION | El pegado mueve goldens Julio Q1/Q2 no declarados | CA-C2/C6: solo desplazamientos DECLARADOS; set-diff vacío fuera de la lista |
| R-PARCIAL | R10 ausente o sin las 2 hojas ⇒ salida a medias | D-H1/H2: fail-fast existente (`No se encontró R10…` / `ERR-PLANTILLA`) ANTES de escribir; sin escritura parcial |

---

## §T0. Evidencia medida (el brief la trae medida; H0 la re-verifica con asserts blindados — NO re-medir a mano lo ya medido)

> Orígenes: oficial `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` (hojas medibles → `sheet36.xml` DetRetri2026092, `sheet37.xml` DetValiRetri2026092, `sheet18.xml` CONSOLIDADO) + insumo `Docs/Prueba Septiembre 1/Insumos/R10_Remuneracion_2026091.xlsx`. H0 = fixture `PegadoR10T0Plan62Tests` (BCL puro `ZipArchive` + `X.Linq` para la plantilla/oficial/manual + ExcelDataReader solo-lectura sobre COPIAS para el R10; rutas explícitas `Docs/Prueba Septiembre 1/...`; prohibido `Insumos.Raiz()` en lo nuevo). Cero código de producción en H0.

### T0a. R10 2026091: exactamente 2 hojas con el mismo layout que la plantilla (asserts T0-01…T0-04)

- **T0-01 (hojas):** el R10 trae EXACTAMENTE 2 hojas: `DetRetri2026091` y `DetValiRetri2026091` (nombres por `NombresHojaPeriodo` del período 2026091). Assert: conteo de `sheet` = 2 + nombres exactos; tercera hoja = FAIL.
- **T0-02 (encabezado corporativo común):** ambas hojas comparten fila 3 `PROCERASEO S.A.S.`/NIT/RMC 2026091, fila 6 Fecha de Proceso 17/09/2026, fila 7 Hora 08:44 AM, rango 01/09–15/09, encabezado de tabla en fila 8, ASE 1..5 en filas 9–13, fila 14 = `Total`. Assert: literales de A3:C8 + B9:B13 = 1..5 + C14 = `Total`.
- **T0-03 (montos como TEXTO monetario):** `"$17.450.250.543"`, negativos `-$5.093.257.797` (misma convención que `ExcelDataReaderDetRetriR10Reader.ParsearNumeroMonetario`). Assert: el lector de bloque parsea ≥1 positivo con `$` y ≥1 negativo con `-$` al centavo; el parseo vive en UNA fuente (`ParseoMonetarioR10`), no duplicado.
- **T0-04 (columna A vacía; B = ASE):** en el R10 la columna A está vacía y la columna B es el id 1..5 (C=CONCESIONARIOS, D..O = montos). Es el MISMO layout que la plantilla DetRetri (sheet36): mismos headers fila 8 (14 columnas), mismos rótulos C9:C13 (incluye ASE3 sin punto final, H-59-4), `D14`/fila 14 `Total`. Assert: `B9:B13 = {1..5}` en R10 y en plantilla; `C9:C13` iguales R10-vs-plantilla (texto exacto).

### T0b. Plantilla DetRetri sheet36: top-block values-only, fórmulas bajo la fila 19 (asserts T0-05…T0-06)

- **T0-05 (slots literales):** filas 3–19 = literales+rótulos (T2 `REPORTE R10- RECAUDO CONSOLIDADO`, F3..F19 literales, M3=2026092, D6/D7 sello, headers fila 8, B9:C13 rótulos, B14/C14=`Total`) y los slots de montos D9:O14 son literales `<v>0</v>` (Plan 59 §3.4, destinos de escritura). Assert: 0 celdas con `<f>` en A3:O19 (barrido XML).
- **T0-06 (fórmulas fuera del pegado):** NO hay `<f>` en filas 3–19 de DetRetri; las fórmulas D23:D28/D32:D36 + auto-chequeos L32:O36 están BAJO la fila 19. Assert: el pegado A3:O19 no toca ninguna `<f>` (diff de `<f>` antes/después = ∅).

### T0c. Plantilla DetValiRetri sheet37: idem + diferencias que consumen lo pegado (asserts T0-07…T0-09)

- **T0-07 (slots literales):** top-block con slots literales en D9:O14. Assert: 0 `<f>` en A3:O14.
- **T0-08 (fórmulas 16–20 intactas y consumidoras):** `D16..O20 = D9..'CONSOLIDADO_TOTAL RECAUDO'!U104..AG104` por ASE (filas 16–20 ↔ filas 9–13) quedan FUERA del pegado y consumen lo pegado. Assert: texto-`<f>` de D16:O20 byte-idéntico antes/después + referencia a su fila 9–13.
- **T0-09 (J9:J14 entra al pegado):** `J9:J14` (`AJUSTE A LA DECENA`) está dentro del rango ⇒ la exclusión de comparador `DetValiRetri!J` (SALE T0d del Plan 29) se RETIRA. Assert: J9:J14 del manual ≠ 0 y J9:J14 del R10 ≠ 0 (hay dato real que pegar, no ceros).

### T0d. CONSOLIDADO sheet18 filas 141–147: estado actual y remapeo exigido (asserts T0-10…T0-12)

- **T0-10 (geometría):** fila 141 = headers; filas 142–146 por ASE (B=1..5); fila 147 = Total (`C147` comparte s-string con `Total`). Assert: B142:B146 = 1..5 + C147 = `Total`.
- **T0-11 (fórmulas por fila k, estado actual):** `D{k} = D104+N104+D125+E125` (internas, sin quincena); `E{k} = 'Recaudo EAAB Reciprocidad'!F21+'Recaudo EAAB + Ciud Limp'!F21+'Recaudo Directa Occidente'!F21+'Recaudo ENEL'!F21+'Recaudo ENERBIT'!F21` con E142→21 … E146→25; `F..J,H,I` internas; `L{k} = D-E-F-G-J-H-I`, `M{k} = INT(ABS(L))=0`, fila 147 SUM. Assert: texto-`<f>` de E142:E146 = Σ con columna F y filas 21:25 (forma exacta citada).
- **T0-12 (desalineación Q1):** esa columna E NO está en la familia E1-4 ⇒ en Q1 queda en F mientras los `Recaudo *` reciben columna D. Assert: en una corrida Q1 real, `E142` referencia `F{21}` mientras el writer escribe `D{21}` (desalineación demostrada, no declarada).

**Puerta H0→H1:** T0-01…T0-12 verdes + doc `Docs/T0-Plan62-Pegado-R10.md`. Si un número del §T0 difiere ⇒ STOP (se re-mide, no se "ajusta" el plan en silencio).

---

## §Doctrina (cambios normativos que este plan introduce)

### D-G3 — El R10 pasa de oráculo a FUENTE de DetRetri/DetValiRetri (REEMPLAZA a G3-D1)

**Texto vigente a retirar** (project-context §Datos Reales): *"R10_Remuneracion_AAAAMMQ.xlsx … Oráculo de VALIDACIÓN del DetRetri calculado (G3) en AMBAS quincenas, NUNCA se escribe."*

**Texto de reescritura para `project-context.md`:**

> `{periodo}/R10_Remuneracion_AAAAMMQ.xlsx` — insumo de PERÍODO con 2 hojas (`DetRetri{AAAAMMQ}` filas 3–19 cols A–O, `DetValiRetri{AAAAMMQ}` filas 3–14 cols A–O). Es la **FUENTE** del pegado de esas hojas en TODOS los períodos y AMBAS quincenas (Plan 62, N1: reemplaza a G3-D1): la app copia el bloque celda-a-celda (valores/texto, estilos de destino preservados) y lo verifica por read-back. El gate total calculado-vs-R10 queda obsoleto para la hoja (tautología); viven el gate per-ASE (o su equivalente DetValiRetri D16:D20 ≈ 0 ±0.5) y la verificación de integridad del pegado.

### D-59H4 — El asiento DetRetri del Plan 59 H4 SALE (lo cubre el pegado)

`EscribirCeldasDetRetri` (D9:D13/D14) + `EscribirDesgloseTrazableDetRetri` (J/L/Cierre + rótulos `Costo de la Interventoría`/`Validación retribución` en `OpenXmlPlantillaWriter.cs:1748-1835`) + sus `Verificar*` dejan de correr: todas sus celdas caen dentro de D9:O13 / J9:J14 ⇒ sobre-escritas por el pegado. El retiro exige traza de consumidores (T2.5): si un consumidor (REMUNERACION_*, INTERVENTORIA, gates) necesita el valor, el valor sigue en el dominio (Inputs), no en la hoja.

### D-61R10 — El gate total calculado-vs-R10 se SUSTITUYE por verificación de integridad (REEMPLAZA al frente R10 del Plan 61 para la hoja)

- **Conservado:** gate per-ASE contra R10 (o equivalente: `DetValiRetri!D16:D20 ≈ 0 ±0.5`, que ya existe y consume lo pegado).
- **Sustituido:** el gate de TOTAL por **verificación de integridad del pegado**: read-back del rango escrito ≡ lectura del R10 (celda a celda, ±0.5 en montos / igualdad de texto en rótulos); si difiere ⇒ fail-fast que nombra hoja+celda+valor esperado vs escrito. El `-1` residual vive donde ya lo tiene el manual (`DetValiRetri!D21`, divergencia documentada excluida del gate — D6); NO se abre categoría nueva.
- **Efecto en Plan 61:** `ClasificadorNovedadesEntrega.DesdeR10` deja de emitir la novedad de TOTAL para la hoja (la hoja es el R10 por construcción); las novedades per-ASE siguen su política de entrega con novedades (H3 del Plan 61 intacto).

### D-31SELLO — El sello de proceso Plan 31 en D6/D7 SALE (lo cubre el pegado)

`SellarProcesoEnDetRetri` (D6/D7 en ambas hojas) se retira: el pegado A3:O19/A3:O14 ya trae Fecha de Proceso 17/09/2026 + Hora 08:44 AM del R10 por construcción. La lectura de `FechaProceso`/`HoraProceso` del R10 y su fail-fast `ERR-FORMATO-FUENTE` se conservan (siguen validando que el R10 traiga el sello legible); lo que sale es la ESCRITURA separada.

### D-E1-4EXT — La familia E1-4 se extiende al bloque 142–147 (Plan 40 + este plan)

Allow-list Phase 0 vigente: 68 celdas (Plan 40). Este plan añade **5 celdas**: `CONSOLIDADO_TOTAL RECAUDO!E142:E146` (Σ `Recaudo *` por ASE, filas 21:25, columna por quincena). `ValidadorBasePeriodo` exige la variante del run y rechaza tercera forma con `ERR-PLANTILLA`.

### D-J9 — `DetValiRetri!J9:J14` vuelve al comparador (RETIRO de la exclusión SALE del Plan 29)

La categoría SALE (manual-externo, `AJUSTE A LA DECENA`) existía porque la app dejaba J9 en 0. Con el pegado, J9:J14 trae dato del R10 ⇒ la exclusión se RETIRA del catálogo versionado (con cita a T0-09).

---

## §2. DESIGN

### D-H0. Método T0 por bytes (read-only bloqueante, cero código de producción)

Forma normativa: BCL puro (`ZipArchive` + `X.Linq`) para plantilla/oficial/manual/administrativo + ExcelDataReader solo-lectura sobre COPIAS para el R10; rutas explícitas a `Docs/Prueba Septiembre 1/Insumos`, sus `Resultados/Administrativo` y `Docs/Prueba Julio-1`, `Docs/Prueba Septiembre-2` (+ `Docs/Insumos` y T0-Plan57 para rutas R10 según el brief). Prohibido `Insumos.Raiz()` en lo nuevo. Cero código de producción en H0: solo fixture `PegadoR10T0Plan62Tests` + doc `Docs/T0-Plan62-Pegado-R10.md`. Doce asserts (T0-01…T0-12).

| Candidata | Veredicto | Motivo |
|---|---|---|
| **M-A: pegado de bloque R10 + extensión E1-4 + reformulación de gates + reclasificación** | **Elegida** | Única que cierra N1 y N2 con la salida indistinguible del manual |
| M-B: recalcular DetRetri bottom-up y "tolerar" el −1 del total | Descartada | El negocio ordenó el copiado; tolerar escondería la divergencia de composición en vez de eliminarla |
| M-C: remapear E142:E146 por contenido (detectar D o F escritas) | Descartada | Viola D-H (dominio, nunca contenido); la columna la decide `NumeroQuincena` |
| M-D: pegar solo D9:D14 (montos) y conservar escrituras de rótulos/sello | Descartada | Deja doble-escritura y estados intermedios; el manual trae el bloque completo y el pegado es celda-a-celda |

### D-H1. Lector de bloque R10 (decisión)

1. **Superficie:** se EXTIENDE `IDetRetriR10Reader` (no interface nueva): `LeerBloqueDetRetri(periodo, rutaR10) → BloqueHojaR10` (rango A3:O19) y `LeerBloqueDetValiRetri(periodo, rutaR10) → BloqueHojaR10` (rango A3:O14). Hojas por `NombresHojaPeriodo` (dominio, nunca enumeración). R10 ausente o sin las 2 hojas ⇒ fail-fast existente (`No se encontró R10…` / `ERR-PLANTILLA`), sin escritura parcial.
2. **Localización por etiquetas:** header fila 8 (`B=ASE`, `D=RECAUDO TOTAL…`, mismo predicado del lector actual); filas 9–13 por `B ∈ 1..5`; fila 14 por `C=Total`; pie DetRetri 15–19 por posición relativa al Total (literales `* Costos de Interventoría…`, `Firma autorizada`). La letra física (B, D..O) NUNCA se nombra en producción salvo el rango del contrato T0 (A:O × filas, geometría citada).
3. **Valores:** el conversor monetario (`$17.450.250.543`, `-$5.093.257.797`) se EXTRAE de `ExcelDataReaderDetRetriR10Reader` a `ParseoMonetarioR10` (Core o Infrastructure-adjacent puro, sin I/O) como fuente ÚNICA; el lector actual lo consume (sin duplicar). Cada celda del bloque = `{ Fila, Columna, Texto, Numerico? }`: montos → numérico redondeado a entero (regla CONSOLIDADO DetRetri: round-half-away); rótulos/fechas/hora → texto exacto (sello D6/D7 como texto, doctrina Plan 31 R-S-1).
4. **Modelo:** `BloqueHojaR10 { Hoja, Periodo, Celdas: IReadOnlyList<CeldaBloqueR10> }` en Core (puro, sin I/O). Sin ramas por período/quincena/ASE (D-H): el rango es del contrato T0, los NOMBRES de hoja son del dominio.

### D-H2. Writer de pegado OpenXML (decisión)

1. **Punto único:** `OpenXmlPlantillaWriter.PegarBloqueR10(workbookPart, periodo, bloqueDet, bloqueVali)` (nombres por `NombresHojaPeriodo`). Corre en la MISMA pasada atómica del writer, DESPUÉS del cálculo CONSOLIDADO (U104:AG108 disponible para las D16:D20 que consumen lo pegado — el pegado escribe las ENTRADAS 9–13, no las fórmulas 16–20) y DESPUÉS de Phase 0 (renombres), y ANTES de las lecturas-oráculo/gates de validación (`ValidacionOracleReader`/`ValidadorBasePeriodo` evalúan D16:D20/D24:D29 desde lo escrito — `ResolutorValorValidacion`). Cota a verificar contra `ProcesadorPeriodo.Ejecutar` antes de fijar (T2.1).
2. **Values-only estricto:** por celda del bloque: si el destino trae `<f>` ⇒ `ERR-PLANTILLA` nombrando hoja+celda (guard pre-escritura, misma familia que `GuardEscrituraFormula`/`EscribirValorNumerico`); si no, se escribe el `<v>`/texto preservando el estilo (`s`) del destino. Solo valores/texto; sin insert/delete (las filas ya existen). Diff de `<f>` antes/después = ∅ (assert H2).
3. **Retiros con traza (T2.5, antes de borrar):** `EscribirCeldasDetRetri` + `VerificarAsientoDetRetri` (+ `VerificarDesgloseTrazableDetRetri`), `EscribirDesgloseTrazableDetRetri` (rótulos `:1819-1835`), sello `SellarProcesoEnDetRetri` (D6/D7), y la escritura "23 celdas" (`:928`). Cada retiro cita sus consumidores (REMUNERACION_*, INTERVENTORIA, gates, comparador) y demuestra que el valor sigue en el dominio o que el consumidor lee lo pegado.
4. **Fail-fast sin parcial:** cualquier celda del bloque ilegible/ausente en el R10 ⇒ `ERR-FORMATO-FUENTE` (período+archivo+hoja+celda) ANTES de escribir la primera celda del bloque.

### D-H3. Reformulación de gates (decisión)

1. **Per-ASE conservado:** `ValidarDetRetriContraR10` por ASE (o su equivalente workbook `DetValiRetri!D16:D20 ≈ 0 ±0.5`) sigue corriendo post-pegado contra `ValidacionOracleReader` (read-back de lo escrito, G-WORKBOOK).
2. **Total sustituido:** el gate de TOTAL calculado-vs-R10 se RETIRA para la hoja y se SUSTITUYE por `VerificarIntegridadPegadoR10`: read-back del rango escrito (A3:O19 / A3:O14) ≡ lectura del R10 (montos ±0.5, textos exactos); divergencia ⇒ fail-fast con hoja+celda+esperado+escrito. Con el pegado, `D14` entregado ≡ R10 por construcción ⇒ la novedad `-1` DESAPARECE.
3. **`D21` residual:** la divergencia de redondeo de composición vive donde ya la tiene el manual (`DetValiRetri!D21`); divergencia documentada excluida del gate (D6). NO se abre categoría nueva.
4. **Plan 61 intacto:** la política de entrega con novedades (H3/H4 del Plan 61: orden escritura→validación, `ClasificadorNovedadesEntrega`, TXT, códigos) NO cambia; solo deja de emitirse la novedad de TOTAL-R10 para la hoja.

### D-H4. Extensión E1-4 al bloque 142–147 (decisión)

1. **Tabla nueva, misma mecánica:** `CeldasRecaudoValidacionConsolidado142 = [(E142,21),(E143,22),(E144,23),(E145,24),(E146,25)]` en `PreparadorBasePipeline` (geometría T0-10/T0-11) + `AplicarVarianteRecaudoValidacionConsolidado142` (reutiliza `ColumnaRecaudoDe(periodo)`: Q1→D, Q2→F; reutiliza `AsegurarVarianteRecaudo`/`ReconstruirFormulaRecaudo` o su generalización sin duplicar) + verificación base en `ValidadorBasePeriodo` (variante del run; tercera forma = `ERR-PLANTILLA`).
2. **Alcance:** SOLO `E142:E146`. `D142:D147`, `F..M`, fila 147 y el resto de CONSOLIDADO intactos (recalculan solos). Las filas 21:25 son la geometría DE LA PLANTILLA (doctrina Plan 40); el `D24` del ejemplo del negocio describe su manual de Q1, no la plantilla.
3. **Allow-list:** 68 → 73 celdas Phase 0. El grep de cierre distingue Phase 0 (`PreparadorBasePipeline`) del procesamiento values-only.

### D-H5. Reclasificación tests/comparador + doctrina (decisión)

1. **Comparador:** la exclusión `DetValiRetri!J` (SALE) se RETIRA del catálogo versionado (T0-09 la invalida: J9:J14 trae dato pegado). Las exclusiones V3 (otra quincena), metadatos (`N3`/nombres salvo proceso) y D6 (`D21`) intactas.
2. **Asserts "Q1 fórmulas protegidas intactas" filas 9–14 → doctrina nueva:** los tests que asertaban que DetRetri-Q1 conserva sus fórmulas en filas 9–14 (protegidas HU-07) se reclasifican: desde este plan, AMBAS quincenas pegan el bloque (la hoja Q1 ya no conserva fórmulas en 9–14 por diseño de negocio).
3. **Doctrina y manual:** `project-context.md` (textos §Doctrina D-G3/D-59H4/D-61R10/D-31SELLO/D-E1-4EXT/D-J9) + manual de usuario (`DetRetri` ya no "Q1 solo valida": AMBAS quincenas escriben desde el R10).

---

## §3. SPEC — Tasks trazadas con AC

### H0 — T0 read-only (bloqueante)

| Task | AC |
|---|---|
| T0.1 Fixture `PegadoR10T0Plan62Tests` (T0-01…T0-04: 2 hojas, encabezado corporativo, texto monetario, layout B=ASE) | 4/4 verdes; rutas explícitas; cero producción |
| T0.2 Asserts plantilla DetRetri/DetValiRetri (T0-05…T0-09: 0 `<f>` en A3:O19/A3:O14; D16:O20 consumidoras intactas; J9:J14 con dato) | 5/5 verdes |
| T0.3 Asserts CONSOLIDADO 141–147 (T0-10…T0-12: geometría, forma E142:E146 con F, desalineación Q1) | 3/3 verdes |
| T0.4 Doc `Docs/T0-Plan62-Pegado-R10.md` (matriz R10-vs-plantilla + forma E142:E146 + veredicto J9) | Doc existe y cita los 12 asserts |

### H1 — Lector de bloque R10

| Task | AC |
|---|---|
| T1.1 Core: `BloqueHojaR10` + `CeldaBloqueR10` (puros) + extensión `IDetRetriR10Reader` (`LeerBloqueDetRetri/LeerBloqueDetValiRetri`) | Compila; 0 warnings; sin I/O en Core |
| T1.2 Extraer `ParseoMonetarioR10` (fuente única) y hacerlo consumir por el lector actual (identidad: D9:D13/D14 byte-idénticos a hoy en 2026091) | `GoldenResumenMesYR10Tests` + `RegresionOficialPlan39Tests` verdes sin tocar |
| T1.3 Infrastructure: `ExcelDataReaderBloqueR10Reader` (o extensión del existente) por etiquetas; fail-fast `ERR-FORMATO-FUENTE`/`ERR-PLANTILLA` con hoja+celda | R10 2026091 → bloques A3:O19/A3:O14 completos; R10 ausente/sin hojas ⇒ fail-fast nombrando archivo+hoja, sin escritura |
| T1.4 Tests `BloqueR10H1Plan62Tests` (per-ASE ±0.5 vs `LeerDetRetri` actual; textos D6/D7 exactos; Total ≡ R10) | Verdes; el lector actual y el de bloque coinciden al centavo |

### H2 — Writer de pegado + retiros

| Task | AC |
|---|---|
| T2.1 Fijar orden en `ProcesadorPeriodo.Ejecutar`: cálculo CONSOLIDADO → pegado → oráculos/gates (cota verificada en código) | Orden citado con archivo:línea; gates leen lo pegado |
| T2.2 `PegarBloqueR10` values-only (estilo destino preservado; guard anti-fórmula pre-celda; diff `<f>` = ∅) | Septiembre-1: DetRetri/DetValiRetri ≡ R10 (read-back); build 0 warnings |
| T2.3 Read-back de integridad `VerificarIntegridadPegadoR10` (hoja+celda+esperado+escrito en el throw) | Divergencia sintética ⇒ fail-fast que nombra la celda |
| T2.4 Retirar sello D6/D7 Plan 31 (conservar su LECTURA + fail-fast `ERR-FORMATO-FUENTE`) | D6/D7 entregados ≡ R10; sin doble-escritura |
| T2.5 Traza de consumidores y retiro de `EscribirCeldasDetRetri`/`EscribirDesgloseTrazableDetRetri`/`Verificar*` + "23 celdas" (`:928`) | Cada retiro cita consumidor y destino del valor en dominio; 0 escritura muerta (grep de cierre) |

### H3 — Gates

| Task | AC |
|---|---|
| T3.1 Conservar gate per-ASE (o equivalente D16:D20 ≈ 0 ±0.5) post-pegado | Septiembre-1: per-ASE verdes |
| T3.2 Retirar gate total calculado-vs-R10 para la hoja; `ClasificadorNovedadesEntrega.DesdeR10` deja de emitir novedad de TOTAL | Septiembre-1: 0 novedad R10 `-1`; `D21` residual donde lo tiene el manual |
| T3.3 Tests `GatesPegadoH3Plan62Tests` (tautología imposible: gate total ausente; integridad read-back roja ante pegado corrupto) | Verdes |

### H4 — Extensión E1-4 (142–147)

| Task | AC |
|---|---|
| T4.1 `CeldasRecaudoValidacionConsolidado142` + `AplicarVariante…` (reutiliza `ColumnaRecaudoDe`) + `ValidadorBasePeriodo` (variante del run; tercera forma `ERR-PLANTILLA`) | Q1 → D{21:25}, Q2 → F{21:25} en texto-`<f>` |
| T4.2 Tests `RecaudoValidacionConsolidado142H4Plan62Tests` (Q1 D / Q2 F / tercera forma FAIL; resto de la fila intacto: D/L/M/147) | Verdes |
| T4.3 Allow-list 68 → 73 + grep de cierre Phase-0 vs values-only | Grep verde |

### H5 — Reclasificación + doctrina

| Task | AC |
|---|---|
| T5.1 Comparador: retirar exclusión `DetValiRetri!J` (versionada, con cita T0-09) | Septiembre-1: comparador no sale por las 2 hojas |
| T5.2 Reclasificar tests afectados (espectro explícito §8): `AsientoDetRetriEscritoPlan59Tests`, `BaseConsistente`, `ComparadorSalidaVsManual`, `EntregaConNovedadesH3Plan61Tests`, `FNew1`/T0 de validación E1-4, goldens | Suite verde con set-diff declarado (solo desplazamientos del cambio) |
| T5.3 Actualizar `project-context.md` (textos §Doctrina) + manual (`DetRetri` AMBAS quincenas) | Diff citado en el Gate |

### Gate — Cierre

| Task | AC |
|---|---|
| G.1 Corrida 2026091 producción: salida SIN novedad R10 `-1`; DetRetri/DetValiRetri ≡ manual Administrativo (±0.5/texto; comparador sin salir por las 2 hojas) | CA-1 |
| G.2 Julio Q1/Q2 goldens: solo desplazamientos DECLARADOS; cero regresión no explicada | CA-2 |
| G.3 Q1 E142:E146 = Σ `Recaudo *`!D{21:25}; Q2 = F{21:25}; `fullCalcOnLoad` alcanzable (abre sin reparación) | CA-4 |
| G.4 Build 0 warnings; suite verde con set-diff explícito | CA-5/CA-6 |

---

## §4. Criterios de aceptación mínimos (refinados del brief)

- **CA-1 (septiembre-1):** corrida de producción SIN la novedad R10 `-1`; hojas DetRetri/DetValiRetri del entregable ≡ manual Administrativo (valores ±0.5 / texto; comparador sin salir por las 2 hojas).
- **CA-2 (Julio Q1/Q2 goldens):** solo desplazamientos DECLARADOS por el cambio; cero regresión no explicada.
- **CA-3 (CONSOLIDADO):** Q1 E142:E146 = Σ `Recaudo *`!D{21:25}; Q2 = F{21:25}; `fullCalcOnLoad` alcanzable.
- **CA-4 (fail-fast):** R10 ausente o sin las 2 hojas ⇒ fail-fast existente (`No se encontró R10…` / `ERR-PLANTILLA`), sin escritura parcial.
- **CA-5 (calidad):** build 0 warnings; suite verde con espectro de reclasificación explícito.
- **CA-6 (integridad):** read-back del pegado con anchors por hoja a la fuente R10 (T2.3).

---

## §5. Architecture Validation Certificate

- **SOLID:** SRP (lector / pegado / gates / Phase-0 en clases separadas); OCP (nuevo rango por tabla citada, no por ramas; `ColumnaRecaudoDe` reutilizado); LSP/ISP (extensión de `IDetRetriR10Reader`, sin ensanchar implementadores ajenos); DIP (Core define `BloqueHojaR10`, Infrastructure lo materializa).
- **Best Practices:** values-only estricto (diff `<f>` = ∅ salvo las 5 celdas Phase-0 declaradas); guard anti-fórmula como red; localización por etiquetas (D-H: cero literales de período/empresa/canal en producción nueva); prohibido `<v>` stale como oráculo; `NombresHojaPeriodo` para las 4 resoluciones de hoja (2 origen + 2 destino).
- **Performance:** el pegado son ≤ 255 celdas (17×15 + 12×15) por corrida multi-ASE sobre la copia de trabajo (no sobre el oficial); el lector R10 abre el insumo una vez (los dos bloques en una pasada o dos pasadas read-only acotadas); sin impacto en el hot path R1/R2/R4.

---

## §6. Clarification Gate (preguntas al Ingeniero — STOP hasta su ratificación)

1. **D-RET-SELL0:** ¿Se retira el sello Plan 31 en D6/D7 (el pegado lo cubre) o se conserva como redundancia? Recomendación senior: **retirar** (menos escrituras activas; el dato ya es del R10).
2. **D-RET-H4:** ¿Retiro total de `EscribirCeldasDetRetri`/`EscribirDesgloseTrazableDetRetri`/`Verificar*` (con traza T2.5) o conservación tras el pegado? Recomendación: **retirar** (evita escritura muerta y estados intermedios).
3. **D-GATE-TOTAL:** ¿Sustituir el gate total por verificación de integridad read-back (el `-1` desaparece por construcción) o conservarlo como novedad? Recomendación: **sustituir** (conservarlo sería tautología).
4. **D-J9:** ¿Retirar la exclusión `DetValiRetri!J` del comparador (J9:J14 trae dato pegado)? Recomendación: **retirar** con cita T0-09.
5. **D-FILAS-2125:** ¿Confirmado que E142:E146 remapea a filas 21:25 de la plantilla (doctrina Plan 40) y que el D24 del ejemplo del negocio describe su manual, no la plantilla? Recomendación: **sí** (la app escribe el recaudo en sus propias filas).
6. **D-D21:** ¿El `-1` residual en `DetValiRetri!D21` queda como divergencia documentada excluida (D6) sin categoría nueva? Recomendación: **sí**.
7. **D-ALCANCE:** ¿El espectro de tests a reclasificar (§8) es completo o hay goldens adicionales bajo paridad (p. ej. octubre-1)? Recomendación: fijarlo en H0 con el trx de entrada.
8. **D-MANUAL:** ¿Texto del manual para "DetRetri AMBAS quincenas desde R10"? Recomendación: redactarlo en H5 con aprobación del área administrativa.

---

## §7. Riesgos y regresiones (detalle)

| ID | Riesgo / regresión | Disparo | Mitigación + test que lo fija |
|---|---|---|---|
| R-PEGADO-FORMULA | Destino con `<f>` en A3:O19/A3:O14 | Plantilla futura | Guard pre-celda `ERR-PLANTILLA` (T2.2); fixture con `<f>` injertada debe enrojecer |
| R-DOBLE-ESCRITURA | Escrituras Plan 59-H4 / Plan 31 / rótulos bajo el pegado | Olvido de retiro | T2.5 traza + grep de cierre (0 llamadas post-retiro) |
| R-TAUTOLOGIA | Gate total tras el pegado | Conservarlo | T3.2 lo retira; T3.3 fija la ausencia |
| R-Q1-DESALINEADA | E142:E146 queda en F en Q1 | No extender E1-4 | T4.1/T4.2 + G.3 |
| R-J9-FALSO-VERDE | Comparador sigue excluyendo J | No retirar SALE | T5.1; septiembre-1 debe comparar J9:J14 |
| R-Q1-PROTEGIDAS | Tests "Q1 fórmulas 9–14 intactas" en rojo | Cambio de doctrina | T5.2 reclasifica con cita (no se reverdece a ciegas) |
| R-PARCIAL | R10 incompleto ⇒ salida a medias | Insumo roto | Fail-fast pre-escritura (T1.3/T2.4); `PreflightInsumosTests` intacto |
| R-ORACULO-STALE | `<v>` como oráculo | Costumbre | G-WORKBOOK en T3.1/T4.2/G.1 |

**Tests afectados (espectro a reclasificar en H5, lista explícita):** `AsientoDetRetriEscritoPlan59Tests` (el asiento sale → pegado; sus asserts de D9:D14 pasan a read-back vs R10), `BaseConsistente` (variante E142:E146 por quincena), `ComparadorSalidaVsManual` (+`ComparadorSalidaVsManualTests`: exclusión J9 OUT), `EntregaConNovedadesH3Plan61Tests` (novedad TOTAL-R10 ya no se emite para la hoja), `FNew1`/T0 de validación E1-4 (allow-list 68→73; `DetRetri!L32:L36/O32:O36` y `DetValiRetri!I24:I29` siguen evaluando lo pegado — verificar que el pegado no mueve sus entradas), goldens Julio Q1/Q2 (desplazamientos declarados: DetRetri/DetValiRetri ahora = R10 del período; E142:E146 por quincena).

---

## §8. Traza de decisiones arbitrales (elecciones propias con fundamento — para aprobación del Ingeniero)

| # | Decisión | Fundamento | Alternativa descartada |
|---|---|---|---|
| D-RET-SELL0 | Retirar sello D6/D7 Plan 31 (conservar lectura+fail-fast) | El pegado trae D6/D7 del R10 por construcción; doble-escritura = estados intermedios confusos | Conservarlo "por si acaso" (escritura muerta) |
| D-RET-H4 | Retirar asiento H4 Plan 59 + rótulos + "23 celdas" (con traza T2.5) | Todas caen en D9:O13/J9:J14 ⇒ sobre-escritas; el valor vive en el dominio | Reclasificarlas como "segunda pasada" (doble fuente de verdad) |
| D-EXTIENDE-R10 | Extender `IDetRetriR10Reader` en vez de interface nueva | El lector ya es la costura R10 (header dinámico + parser monetario + fail-fast); nueva interface duplicaría localización y mensajes | `IBloqueR10Reader` separada (dos localizaciones del mismo header) |
| D-PARSE-UNICO | Extraer `ParseoMonetarioR10` como fuente única | El brief lo exige (no duplicar); el lector actual lo consume sin cambiar D9:D13/D14 | Copiar el conversor al lector de bloque (divergencia futura) |
| D-GATE-TOTAL | Sustituir gate total por integridad read-back | Con el pegado, D14 ≡ R10 por construcción: comparar el computado sería tautología; lo que prueba es que el pegado quedó (read-back) | Conservar el gate (falso verde permanente) |
| D-PERASE-CONSERVA | Conservar gate per-ASE (o D16:D20 ≈ 0) | Sigue probando que lo pegado cuadra por ASE contra el R10/consumidores; ya existe y consume lo escrito | Retirar todos los gates R10 (quedarse sin malla) |
| D-FILAS-2125 | Remapeo a filas 21:25 (plantilla), no D24 del ejemplo | Doctrina Plan 40: la app escribe el recaudo en sus propias filas; el negocio describe su manual | Seguir el D24 literal (rompería la geometría del template) |
| D-J9-OUT | Retirar exclusión `DetValiRetri!J` | T0-09: J9:J14 trae dato real del R10; la exclusión nació del 0 de la app, no del negocio | Conservarla (falso verde en AJUSTE A LA DECENA) |
| D-ORDEN | Pegado tras CONSOLIDADO+Phase 0 y antes de oráculos/gates | Las D16:D20 consumen U104:AG108 + lo pegado; los gates evalúan lo escrito (`ResolutorValorValidacion`) | Pegar al final (los gates leerían ceros) o al inicio (Phase 0 renombraría sobre lo escrito) |
| D-D21-D6 | `-1` residual en D21 como divergencia documentada (D6) | Ya vive donde lo tiene el manual; abrir categoría nueva sería inflar taxonomía | Nueva categoría Plan 62 (ruido) o tolerarlo en silencio (falso verde) |

---

## §9. Fuera de alcance (vinculante)

Cualquier otra hoja del workbook; cambios de negocio en fórmulas fuera de `E142:E146`; deuda `R-SLOT-Q1-ENERBIT` (Planes 41/46); F-NEW-1 hardcode M; RecaudosReversados; re-congelamientos de mapas fuera del espectro §7/§8; commits/push/checkout/reset.

## §11. Registro de ratificación (2026-10-29, fixture `PegadoR10T0Plan62Tests` 12/12 + `Docs/T0-Plan62-Pegado-R10.md`)

El Ingeniero ratificó las **5 precisiones** del byte vs. §T0 y emitió UNA aclaración vinculante:

| # | Precisión ratificada | Ajuste normativo |
|---|---|---|
| R-1 | El negativo real del R10 2026091 es `DetRetri2026091!J10 = -$5.093.259.797` (no `257`) | D-H1/D-PARSE-UNICO citan ese valor |
| R-2 | Los literales corporativos viven en `F3/H3/C6/D6/C7/D7/F7/G7/I7/J7 + L3/M3 (DetRetri) / J3/K3 (DetValiRetri)`; `A3:B8` y `E3:E8` vacías | T0-02/lectura por etiquetas usan esas coordenadas como CITA, no como contrato |
| R-3 | `C147 = 'TOTAL'` (s-string distinto del `Total` de `DetRetri!C14`) | la resolución por etiqueta de la fila de cierre del bloque 142–147 es case-insensitive / por s-string real; prohibido igualdad exacta con `'Total'` |
| R-4 | El rótulo real es `F6 = R10 - Recaudo Consolidado` (no `REPORTE R10- RECAUDO CONSOLIDADO`) | el literal T2 de §T0/T0-05 se corrige en el doc |
| R-5 | `L145` es fórmula máster compartida (`si=68`) y `M142` máster (`si=67`; `M143:M147` seguidoras) | **H4 OBLIGACIÓN**: al reescribir `E142:E146` se debe preservar la geometría de fórmulas compartidas de la fila (`si`/`ref`); gate en T4.2 |
| R-Aclareo | **Los encabezados por fuera (filas 3–8: PROCERASEO/NIT/Código Remuneración del período/Fecha de Proceso/Hora/Fecha Desde-Hasta) VIAJAN pegados desde el R10 del período que se ejecuta** — no se conservan los del template ni se reescriben aparte. El pegado A3:O19/A3:O14 ES el transportador de esos encabezados (ratificación explícita del Ingeniero) | confirma D-H2 (values-only sobre el rango completo) y hace a D-31SELLO (retiro del sello)binding: el sello D6/D7 del R10 llega pegado |

**Riesgo operativo añadido a H1 (T1.2):** `ParsearNumeroMonetario` documenta "',' decimal" pero implementa `Where(char.IsDigit)` ⇒ `"$1.234,56"` se leería `123456`. La extracción a `ParseoMonetarioR10` debe resolver DOCUMENTO-vs-CÓDIGO con test dedicado (decidida aquí: atributo del fixture — los montos 2026091 son enteros, pero el fix es precondición del retiro del gate con valores decimales en períodos futuros).

## §12. Ratificaciones H1 y refinamientos del code-reviewer (2026-10-29)

El Ingeniero ratificó las 4 decisiones del implementer (predicado `B='ASE' ∧ C='CONCESIONARIOS'`; montos→`Numerico` sin redondeo / textos exactos / vacías presentes; bloque estricto sin parcial; `ParseoMonetarioR10` en Infrastructure). El review (PASS-with-notas) añade obligaciones trazadas:

| ID | Refinamiento | Dónde aplica |
|---|---|---|
| N-1 | **Hardening del signo en `ParseoMonetarioR10`** (hallazgo #2, WARNING): el signo `-` solo es válido en posición inicial (tras el `$`); textos no monetarios con guion (p. ej. separadores de fecha) se RECHAZAN en `TryParseTexto` (false), jamás se adivina negativo. Con test adversarial | `ParseoMonetarioR10` |
| N-5 | **Doble estándar prohibido**: si la lectura conserva el monto SIN redondear, el pegado de H2 TAMBIÉN escribe el valor sin redondear (condición del read-back ±0.5; hallazgo #1) | T2.2 |
| R5-GATE | Preservar la geometría de fórmulas compartidas (`si`/`ref`) bajo/near las egiones tocadas; T4.2 gatea el máster `L145`/`M142` | T2.2/T4.2 |
| H3-SEMANTIC | El read-back de H3 compara por VALOR semántico (numérico/±0.5 en montos; texto exacto en rótulos; forma-agnóstico ante `B9:B13` texto-vs-numérico de un R10 futuro) | T2.3 |
| D-DEFENSIVA (sugerencia, NO obligatoria) | "pin fila 8" en `LeerDetRetri` simétrico al del bloque — delegado a decisión futura, no en H2 | — |

---

## §10. Entregables y reporte

- **Ruta del plan (este documento):** `plans/62 - pegado-r10-det-detvali-y-validacion-recaudo-q1.md`.
- **Al aprobar:** H0 produce `Docs/T0-Plan62-Pegado-R10.md` + fixture `PegadoR10T0Plan62Tests` (23+ asserts T0-01…T0-12); H1–H5 producen código + tests citados; el Gate produce trx + set-diff + corridas 2026091/2026071/2026072.
- **Fe de errata (H5, T5.5):** la cabecera `gate_entrada.plantilla_oficial` citaba la plantilla v1 (`Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`); el pipeline consume **`Insumos.PlantillaOficialV4`** = `Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx` (VIGENTE). La geometría medida en H4 (`E142:E146`, `L145` máster `si=68`, `M142` máster `si=67`) es **idéntica v1-v4** ⇒ no invalida ningún assert T0 ni T4.
- **Estado:** **CERRADO salvo Gate** — H0…H5 ejecutadas y verificadas (H5: `ComparadorPegadoH5Plan62Tests`, reclasificación de fixtures del data drift, doctrina y manuales). El cierre formal (corridas 2026091/2026071/2026072 + set-diff de la suite completa) queda en el **Gate**, que corre en sesión aparte.
