# T0 — Plan 53: contrato celda → sección esperada de los operandos empresa R1

| Campo | Valor |
|---|---|
| Plan | `plans/53 - contrato-seccion-por-operando-r1.md` (**APROBADO 2026-10-22**, §12 ratificado 7/7) |
| Fase | **H0 (read-only, bloqueante)** — gate `H0 → H1` |
| Fecha | 2026-10-22 |
| Veredicto | **VERDE con 1 desviación declarada de evidencia (§6) y 2 Δ adicionales a la tabla §9 (§5.3)** — H1 habilitado |
| Alcance ejecutado | Cero código de producción. Solo lectura BCL (ZIP+XML) de las 20 fuentes R1 reales (`Docs/Prueba {Julio-1, Julio-2, Agosto-2, Octubre-1}/Insumos/{1-Promoambiental…5-Área Limpia}/Recaudoporcomponente_*.xlsx`, hoja `Sheet1`) + oráculo manual (`REMUNERACION_*!D{8+ase}`) + plantilla oficial v2 (solo lectura) |
| Invariantes verificadas antes/después | v2 `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` · v1 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` (intactas) |
| Fixture | `Remuneracion.IntegrationTests/SeccionPorOperandoR1T0Tests.cs` (**11 verdes**, read-only, BCL) |

---

## 0. Resumen ejecutivo

| Gate | Veredicto | Evidencia |
|---|---|---|
| **S1.1** matriz 52 grupos / 180 filas-empresa, firma, cardinalidades, partición/orden | **VERDE** | §2.1–§2.2 (180/180 tipadas; 0 empates; 0 intercalados; 0 formas de 4+; A/B/E vacíos en toda fila-empresa) |
| **S1.2** tabla celda → sección esperada (**56** entradas) + evidencia de fórmula consumidora | **VERDE con desviación declarada** | §3 (tabla) + §6 (el `<f>` consumidor solo existe en 6 celdas de la v2; el resto son slots de valor) |
| **S1.3** casos cardinales | **VERDE** | §4 |
| **S1.4** inventario R2/R4 (D6) | **VERDE** | §5.4 (R2: 50 pares con 2 filas + 2 pares con 0; **0 pares con 1 fila**; R4: ≤ 1 fila siempre) |
| **S1.5** lista cerrada de Δ | **VERDE (17 filas §9 confirmadas + 2 añadidas con evidencia)** | §5.1–§5.3 |
| **S1.6** tipabilidad (`SeccionPorCierre`) | **VERDE** | §2.1 (180/180, 0 `null`, 0 empates) |
| **S1.7** arbitraje del oráculo manual | **VERDE (4/4 al centavo a favor de la sección declarada)** | §5.5 |

**Lectura:** el contrato por celda reproduce exactamente la lista §9 del plan (12 Δ + 5 identidades) y **corrige** el gate Σ en 2 combinaciones (ago-2 ASE2 y ago-2 ASE5 pasan de fallar a cerrar), sin empeorar ninguna (§5.6). Los dos Δ adicionales (§5.3) son consecuencia **normativa** de D5 (misma regla que §9 #1/#2/#5/#8/#12) y el propio §9 autoriza al T0 a ampliar la lista antes de H1.

---

## 1. Método (reproducible, read-only)

1. **Ventana explícita por etiquetas**: primera fila con `E` normalizado = `vlrservicio` hasta la fila invariante `A='Total' ∧ B` vacío; filas fantasma (sin texto ni valor) excluidas; header detectado como la fila anterior no vacía (se asertó `Total` en la columna **F = índice 5**, el contrato del mapa).
2. **Firma 1:1**: cada fila se reconstruye como `FilaEspejoR1` (A–E + `Total`=F + `SERVICIO ESPECIALES`) y se tipa con el predicado vigente `ContratoSeccionPeriodFreeR1.SeccionPorCierre` (Plan 46 §2.2) y, en contraste, `SeccionEspejoR1.PorFronteraArriba` (Plan 41/45).
3. **Filas-empresa**: `C = label(empresa, ase)` (por `SinonimosEmpresaR1.SonMismaEmpresa`, catálogo Plan 34) `∧ D='Total'`, separando `i < indiceComponente` (Totales) de `i > indiceComponente` (Subs).
4. **Resolución VIEJA** (la del código actual, `ExcelDataReaderWorkbookLeafInputReader.cs:1585-1637`): `OrderByDescending(|F|)` → `Main=[0]`, `Medio=[1]`, `Menor=[count-1]`.
5. **Resolución NUEVA** (D1/D4/D5): secciones del bloque **de la empresa** en orden de aparición → `Primera` = la primera **solo si hay ≥2**; `UltimaOportuna`/`UnicaRecaudo` = la última; `Intermedia` = la de en medio **solo si hay exactamente 3**; ausente → **0**; 2 filas en la misma sección → **STOP**.
6. **Oráculo**: `REMUNERACION_<empresa>!D{8+ase}` del manual del administrativo (julio-2 / agosto-2) y el gate Σ del lector de producción (`leaf.R1.TotalOportunoEsperadoPorAse`).
7. Sin escrituras: no se abrió ninguna plantilla en modo escritura, no se generaron salidas, no se tocó el writer.

---

## 2. S1.1 / S1.6 — matriz re-verificada

### 2.1 Gates duros (fixture `S1_1_Matriz…` + `S1_6_Tipabilidad…`)

- **52 grupos** (= 13 pares empresa×ASE del mapa × 4 períodos) → **180 filas-empresa**. 1 grupo con 0 filas: `octubre-1 ASE2 NUEVO ESQUEMA` (ausencia total, Plan 50 A1).
- Cardinalidades de bloque (no vacíos): **1 fila = 42** (14 Totales + 28 Subs) · **2 filas = 42** (19 + 23) · **3 filas = 18** (18 + 0) · **0 filas = 1** (el grupo ausente). **Ninguna forma de 4+.**
- Firma de toda fila-empresa: `A="" B="" C=<label> D="Total" E=""` → **R3 (E no-`Total`) no se dispara**.
- **Tipabilidad 180/180** por `SeccionPorCierre`: 0 `null` (nadie es AFaseo), **0 empates** (dos filas-empresa de la misma empresa en la misma sección) y **0 secciones intercaladas** → R2 y R4-riesgo no se disparan.

### 2.2 Partición y orden: `SeccionPorCierre` vs `PorFronteraArriba`

En **180/180** filas-empresa las dos lecturas devuelven **etiquetas distintas** (la sección se abre por una frontera y se cierra por otra; hasta dos `Oportuno` consecutivos difieren en ordinal) pero inducen **exactamente la misma partición y el mismo orden** (asertado par a par). Como la resolución solo necesita (partición, orden, "primera vs no primera"), el desacuerdo de etiqueta **no altera ningún operando** → se confirma la corrección de la premisa D1-del-52 citando Plan 46 §2.2 (DEFECTO 2). `SeccionPorCierre` es el tipador único.

### 2.3 Marca de "primera sección"

La única marca estable de "primera" es **"no tiene frontera arriba"**, que en el bloque **de la empresa** se traduce a *"la empresa tiene ≥2 secciones propias"*. El **tipo** de la primera sección varía (`Aplicacion#1` en julio-1/2, agosto-2 ASE4/5, octubre-1 ASE1/2/4; `Oportuno#1` en agosto-2 ASE3, octubre-1 ASE3/5) → "primera" **no** puede confundirse con "extemporánea".

---

## 3. S1.2 — Tabla celda → sección esperada (56 entradas)

`SeccionEsperada` por celda del mapa `EditablesR1PorEmpresa` (13 claves × celdas). **El mapa tiene 56 celdas** (no 53: la descomposición del §D2 del plan —`ASE1 3+6, ASE2 3+4+4, ASE3 4+6, ASE4 3+6+6, ASE5 4+3+4`— suma 56; el literal "53" del plan es un error aritmético, la descomposición y el mapa coinciden en 56). El fixture aserta **igualdad exacta de conjuntos** mapa ↔ tabla.

| ASE | Empresa | Celda | Operando | Sección esperada | Evidencia |
|---|---|---|---|---|---|
| 1 | 2 ENEL | F33 | SubsMain | UltimaOportuna | subs ENEL 1 sección (Oportuno#2) |
| 1 | 2 ENEL | F14 | TotalMain | UltimaOportuna | totales ENEL 1 sección (Oportuno#1) |
| 1 | 2 ENEL | L14 | EspecialesMain | UltimaOportuna | misma fila que F14 |
| 1 | 4 OCCIDENTE | F40 | SubsMain | UltimaOportuna | subs OCC 2 secciones → Oportuno#2 |
| 1 | 4 OCCIDENTE | F24 | TotalMain | UltimaOportuna | totales OCC 2 secciones → Oportuno#1 |
| 1 | 4 OCCIDENTE | L24 | EspecialesMain | UltimaOportuna | misma fila que F24 |
| 1 | 4 OCCIDENTE | F29 | SubsMenor | **Primera** | subs OCC 2 secciones → Aplicacion#2 |
| 1 | 4 OCCIDENTE | F9 | TotalMenor | **Primera** | totales OCC 2 secciones → Aplicacion#1 |
| 1 | 4 OCCIDENTE | L9 | EspecialesMenor | **Primera** | misma fila que F9 |
| 2 | 1 RECIP | F102 | TotalMain | UltimaOportuna | totales RECIP 2 secciones → Oportuno#2 |
| 2 | 1 RECIP | F122 | SubsMain | UltimaOportuna | subs RECIP 1 sección |
| 2 | 1 RECIP | F86 | Block5Total | **SinSeccion** | canal `D=5` (semántica intacta) |
| 2 | 2 ENEL | F117 | SubsMain | UltimaOportuna | subs ENEL 1 sección |
| 2 | 2 ENEL | F95 | TotalMain | UltimaOportuna | totales ENEL 2 secciones → Oportuno#2; `<f>` v2 `F95=F69+F46+F15-L15-L46` |
| 2 | 2 ENEL | F83 | TotalMenor | **Primera** | totales ENEL 2 secciones → Oportuno#1 (f7) |
| 2 | 2 ENEL | L95 | EspecialesMain | UltimaOportuna | misma fila que F95 |
| 2 | 4 OCCIDENTE | F129 | SubsMain | UltimaOportuna | subs OCC 1 sección |
| 2 | 4 OCCIDENTE | F112 | TotalMain | UltimaOportuna | totales OCC 2 secciones; consumida por `<f>` v2 `F264=F193+F150+F112--L112-L150` |
| 2 | 4 OCCIDENTE | F89 | TotalMenor | **Primera** | totales OCC 2 secciones → Oportuno#1 |
| 2 | 4 OCCIDENTE | L112 | EspecialesMain | UltimaOportuna | misma fila que F112 |
| 3 | 2 ENEL | F246 | SubsMain | UltimaOportuna | subs ENEL 1 sección |
| 3 | 2 ENEL | F227 | TotalMain | UltimaOportuna | totales ENEL 2 secciones → Oportuno#2 |
| 3 | 2 ENEL | F213 | TotalMenor | **Primera** | totales ENEL 2 secciones → Oportuno#1 |
| 3 | 2 ENEL | L227 | EspecialesMain | UltimaOportuna | misma fila que F227 |
| 3 | 4 OCCIDENTE | F253 | SubsMain | UltimaOportuna | subs OCC 1 sección |
| 3 | 4 OCCIDENTE | F237 | TotalMain | UltimaOportuna | totales OCC 2 secciones → Oportuno#2 (f28 jul-2 / f23 oct-1) |
| 3 | 4 OCCIDENTE | F216 | TotalMenor | **Primera** | totales OCC 2 secciones → Oportuno#1 (f7 jul-2) |
| 3 | 4 OCCIDENTE | L237 | EspecialesMain | UltimaOportuna | misma fila que F237 |
| 3 | 4 OCCIDENTE | F242 | SubsMenor | **Primera** | subs OCC 1 sección → ausente ⇒ 0 |
| 3 | 4 OCCIDENTE | F222 | TotalMedio | **Intermedia** | solo con 3 secciones; con 2 → ausente ⇒ 0 |
| 4 | 1 RECIP | F409 | SubsMain | UltimaOportuna | subs RECIP 2 secciones |
| 4 | 1 RECIP | F380 | TotalMain | UltimaOportuna | totales RECIP 2 secciones (oct-1: 1 sección `Aplicacion#1`) |
| 4 | 1 RECIP | L380 | EspecialesMain | UltimaOportuna | misma fila que F380 |
| 4 | 2 ENEL | F404 | SubsMain | UltimaOportuna | subs ENEL 2 secciones |
| 4 | 2 ENEL | F373 | TotalMain | UltimaOportuna | totales ENEL 3 secciones → Oportuno#2 |
| 4 | 2 ENEL | F353 | TotalMenor | **Primera** | totales ENEL 3 secciones → Oportuno#1; `<f>` v2 `F416=F353`, `F415=F330+F353+F299-L299-L330` |
| 4 | 2 ENEL | L373 | EspecialesMain | UltimaOportuna | misma fila que F373 |
| 4 | 2 ENEL | F395 | SubsMenor | **Primera** | subs ENEL 2 secciones → Aplicacion#2 |
| 4 | 2 ENEL | F362 | TotalMedio | **Intermedia** | totales ENEL 3 secciones → Aplicacion#1 |
| 4 | 4 OCCIDENTE | F416 | TotalMain | UltimaOportuna | totales OCC 3 secciones → Oportuno#2; `F416=F353` (v2) |
| 4 | 4 OCCIDENTE | F390 | TotalMain | UltimaOportuna | totales OCC 3 secciones → Oportuno#2 |
| 4 | 4 OCCIDENTE | F356 | TotalMenor | **Primera** | totales OCC 3 secciones → Oportuno#1 |
| 4 | 4 OCCIDENTE | L390 | EspecialesMain | UltimaOportuna | misma fila que F390 |
| 4 | 4 OCCIDENTE | F400 | SubsMenor | **Primera** | subs OCC 2 secciones → Aplicacion#2 |
| 4 | 4 OCCIDENTE | F368 | TotalMedio | **Intermedia** | totales OCC 3 secciones → Aplicacion#1 |
| 5 | 2 ENEL | F502 | SubsMain | UltimaOportuna | subs ENEL 1 sección |
| 5 | 2 ENEL | F483 | TotalMain | **UnicaRecaudo** | totales ENEL 1 sección (jul-2/ago-2/oct-1) |
| 5 | 2 ENEL | F474 | TotalMenor | **Primera** | totales ENEL: 1 sección jul-2/ago-2/oct-1 ⇒ ausente ⇒ 0 |
| 5 | 2 ENEL | L483 | EspecialesMain | **UnicaRecaudo** | misma fila que F483 |
| 5 | 3 ENERBIT | F505 | SubsMain | UltimaOportuna | subs ENERBIT 1 sección |
| 5 | 3 ENERBIT | F487 | TotalMain | **UnicaRecaudo** | totales ENERBIT 1 sección — **mapa (3,5) sin `TotalMenor`** (prueba de fuego) |
| 5 | 3 ENERBIT | L487 | EspecialesMain | **UnicaRecaudo** | misma fila que F487 |
| 5 | 4 OCCIDENTE | F512 | SubsMain | UltimaOportuna | subs OCC 1 sección |
| 5 | 4 OCCIDENTE | F497 | TotalMain | **UnicaRecaudo** | totales OCC 1 sección (jul-2: Aplicacion#1+Oportuno#1 → última) |
| 5 | 4 OCCIDENTE | F477 | TotalMenor | **Primera** | totales OCC 1 sección jul-2/ago-2/oct-1 ⇒ ausente ⇒ 0 |
| 5 | 4 OCCIDENTE | L497 | EspecialesMain | **UnicaRecaudo** | misma fila que F497 |

**Coherencia rol ↔ operando (asertada):** `TotalMain`/`EspecialesMain`/`SubsMain` → `UltimaOportuna` (ASE5: `UnicaRecaudo`); `TotalMenor`/`EspecialesMenor`/`SubsMenor` → `Primera`; `TotalMedio`/`EspecialesMedio` → `Intermedia`; `Block5Total` → `SinSeccion`.

---

## 4. S1.3 — Casos cardinales (asertados con las fuentes reales)

| Caso | Medición | Contrato |
|---|---|---|
| **Oct-1 ASE5** (3 empresas, 1 fila en la única sección de recaudo) | `F483=7.414.210.944,89` (ENEL f8) · `F487=160.186.886,97` (ENERBIT f12) · `F497=134.612.117,14` (OCC f22), todos en `Oportuno#1` (única sección del bloque de Totales) | `Main(UnicaRecaudo)` = valor; `F474/F477(Primera)` = 0; **ENERBIT sin `TotalMenor` no pierde nada** (el mapa `(3,5)` no lo declara) |
| **Oct-1 ASE3 OCCIDENTE** (solo oportuna) | `f23 Oportuno#2 = 286.711.907,53`; la empresa tiene **1** sección propia (la hoja tiene 2: `Oportuno#1` solo con ENEL) | `Main` = 286.711.907,53; `F216/F242(Primera)` = 0; `F222(Intermedia)` = 0 |
| **Ago-2 ASE2 ENEL/OCCIDENTE** (magnitud ≠ ordinal) | ENEL: `f7 O1 = −1.414.859,14 · f17 A1 = 131.543,19 · f25 O2 = 24.836.019.402,94`; OCC: `f12 O1 = −28.650,84 · f20 A1 = 24.890,00 · f44 O2 = 2.144.338.413,37` | `TotalMenor` = **f7/f12** (primera), `TotalMedio` = **f17/f20** (intermedia) — por **posición de sección**, no por magnitud |
| **Jul-2 ASE5 ENEL** (1 fila Total no-primera) | `f18 Oportuno#1 = 12.823.196.468,26` (la hoja trae `Aplicacion#1` antes) | `Main(UnicaRecaudo)` = valor; `F474(Primera)` = 0 |
| **2 secciones** (19 grupos) | `Intermedia` ausente | `TotalMedio` = 0 |
| **3 secciones** (18 grupos) | `Intermedia` = la de en medio | `TotalMedio` = valor de esa sección (p. ej. jul-1 ASE3 OCCIDENTE = `f16 Aplicacion#1 = 5.698.824,85`) |

---

## 5. S1.5 — Lista cerrada de Δ

### 5.1 Los 12 Δ de la tabla §9 (confirmados al centavo)

| # | Período | ASE | Empresa | Celda | Hoy | Plan 53 | Δ |
|---|---|---|---|---|---|---|---|
| 1 | Jul-2 | 1 | OCCIDENTE | F29 | 168.024.399,69 | 0 | −168.024.399,69 |
| 2 | Jul-2 | 3 | OCCIDENTE | F242 | 404.748.758,01 | 0 | −404.748.758,01 |
| 3 | Jul-2 | 3 | OCCIDENTE | F222 | −16.458,46 | 0 | +16.458,46 |
| 4 | Jul-2 | 5 | ENEL | F474 | 12.823.196.468,26 | 0 | −12.823.196.468,26 |
| 5 | Ago-2 | 1 | OCCIDENTE | F29 | 206.290.801,38 | 0 | −206.290.801,38 |
| 6 | Ago-2 | 2 | ENEL | F83 | 131.543,19 (mag) | −1.414.859,14 (f7) | −1.546.402,33 |
| 7 | Ago-2 | 2 | OCCIDENTE | F89 | 24.890,00 (mag) | −28.650,84 (f12) | −53.540,84 |
| 8 | Ago-2 | 3 | OCCIDENTE | F242 | 360.488.591,23 | 0 | −360.488.591,23 |
| 9 | Ago-2 | 3 | OCCIDENTE | F222 | −522,80 | 0 | +522,80 |
| 10 | Ago-2 | 5 | ENEL | F474 | 13.260.342.962,39 | 0 | −13.260.342.962,39 |
| 11 | Oct-1 | 3 | OCCIDENTE | F216 | 286.711.907,53 | 0 | −286.711.907,53 |
| 12 | Oct-1 | 3 | OCCIDENTE | F242 | 145.489.038,47 | 0 | −145.489.038,47 |

**Identidades blindadas (Δ = 0):** `F483/L483` (ENEL ASE5 oct) · `F487/L487` (ENERBIT ASE5 oct) · `F497/L497` (OCCIDENTE ASE5 oct) → §9 #13–#17 + `L487`.

### 5.2 Δ de comportamiento (hoy throw → 0)

- `Oct-1 ASE3 OCCIDENTE F222` (`TotalMedio`): hoy **lanza** (`ValorDe(filasTotal, 1)` con 1 fila) → 0. Es el residuo `R-PARCIAL-OCT-ASE3-R1-OCC` (Plan 50 §8) que este plan cierra; su test (`S2_1_ParcialConLabelPresente_SigueSiendoFailFast`) se reescribe a aserciones de valor (§9 "baja+alta declarada").

### 5.3 Δ adicionales al §9 (2 filas, consecuencia normativa de D5)

| # | Período | ASE | Empresa | Celda | Hoy | Plan 53 | Δ | Motivo |
|---|---|---|---|---|---|---|---|---|
| 18 | Oct-1 | 5 | ENEL | F474 | 7.414.210.944,89 | 0 | −7.414.210.944,89 | D5: *"`Main` = valor, **`Menor`/`Medio` = 0 donde el mapa los declare**, sin pérdida"* — misma regla que §9 #1/#2/#5/#8/#12 |
| 19 | Oct-1 | 5 | OCCIDENTE | F477 | 134.612.117,14 | 0 | −134.612.117,14 | ídem |

`VisibleR1` resultante (oct-1, sin manual): ENEL `14.647.388.373,02 → 7.233.177.428,13`; OCCIDENTE `306.177.388,95 → 171.565.271,81`; ENERBIT `139.510.580,00 → 139.510.580,00` (**intacto**: su mapa no declara `TotalMenor`). Son la eliminación del **doble conteo** que hoy produce `Main == Menor` en un bloque de una sola sección (el mismo defecto arbitrado por el manual en jul-2/ago-2 ASE5-ENEL).

**Nota:** §9 no enumeraba estas 2 celdas porque su columna "Plan 53" se derivó de la regla D3-del-52 (`única primera ⇒ Menor`), derogada por D5. El propio §9 autoriza la ampliación: *"Si el T0 encontrara una celda fuera de estas 17 con Δ → se añade a la lista **antes** de H1 con evidencia"*.

### 5.4 Inventario R2/R4 (D6) — re-verificado

- **R2** (`B=label ∧ C=Total`): 50 de los 52 pares con **exactamente 2** filas (una en `Componente`, una en `Subs/Cont`); **2** pares con **0** filas (`ENERBIT` ASE5 julio-1 y octubre-1, ausencia total Plan 50 A1); **ningún par con 1 fila** ⇒ el defecto gemelo **no tiene caso real** ⇒ no se requiere plan hijo (ratificación §12-3); el riesgo queda declarado como latente.
- **R4** (`A=label ∧ B=Total`): todos los pares con 0 o 1 fila; **ninguno con 2+** ⇒ `FirstOrDefault` sin índice de ocurrencia ⇒ D6 confirmado.

### 5.5 S1.7 — Arbitraje del manual (4/4 al centavo)

| Caso | Hoy (ocurrencia) | Plan 53 (sección) | Manual | Veredicto |
|---|---|---|---|---|
| Jul-2 ASE5 ENEL | 24.524.354.313,16 | **11.701.157.844,90** | `REMUNERACION_ENEL!D13` = 11.701.157.844,90 | **Plan = manual** |
| Ago-2 ASE5 ENEL | 25.008.738.395,88 | **11.748.395.433,49** | `REMUNERACION_ENEL!D13` = 11.748.395.433,49 | **Plan = manual** |
| Ago-2 ASE2 ENEL | 19.239.204.490,16 | **19.237.658.087,83** | `REMUNERACION_ENEL!D10` = 19.237.658.087,83 | **Plan = manual** |
| Ago-2 ASE2 OCCIDENTE | 2.217.333.974,79 | **2.217.280.433,95** | `REMUNERACION_OCCIDENTE_ Directa!D10` = 2.217.280.433,95 | **Plan = manual** |

**Julio-1 (Q1 con conciliación activa): identidad total, Δ = 0 en los 5 ASE** (43 filas-empresa) — el único bloque donde la identidad es exigible hoy.

### 5.6 Contraste con el gate Σ (Σ visibles-empresa == visible de bloque)

| Combinación | Hoy | Plan 53 |
|---|---|---|
| Ago-2 ASE2 | falla (Δ = 15.6M) | **cierra** |
| Ago-2 ASE5 | falla (Δ = 13.260.342.962,39) | **cierra** |
| Resto de combinaciones evaluables | igual (cerradas o ya fallando por residuos ajenos: `R-SIGMA-Q2`, jul-1 ASE2, jul-2 ASE1/3/5, ago-2 ASE1, oct-1 ASE1/4) | **ninguna pasa de cerrar a fallar** |

Oct-1 ASE3/ASE5 no son evaluables hoy (el bloque ni se calcula: `MapearR1` Q1 exige `Aplicacion`/3 `Mes`) → el gate no se alcanza; el plan no puede empeorarlo.

---

## 6. Desviación declarada de evidencia (D2.1 — geometría del bloque oficial)

**Hecho medido en bytes (v2, `Reporte Componentes R1`):**

1. La plantilla oficial trae, por ASE, un bloque **detalle** (con filas-empresa y fronteras `Mes/Total`, `Aplicacion…/Total`) y un bloque **resumen** (`TOTAL/OPORTUNO`, `<empresa>/OPORTUNO`, `EXTEMPORANEO`, `Subsidio(-)/Contribucion(+)`). Fronteras de bloque medidas: `Total` en 72 / 210 / 355 / 550 / 676; detalle ASE1 6-72, ASE2 110-210, ASE3 289-355, ASE4 450-550, ASE5 603-676 (coincide con `RangoDatosAse` del plan).
2. **Las celdas del mapa `EditablesR1PorEmpresa` NO caen en el bloque detalle del ASE que declaran** en 34 de 56 casos: ASE2 `F83/F86/F89/F95/F102` están en el resumen de ASE1 (75-106); ASE3 `F213…F253` en el resumen de ASE2 (213-285); ASE4 `F356…F416` en el resumen de ASE3 (358-449); ASE5 `F474…F512` en el detalle de ASE4 (450-550). Solo ASE1 (9 celdas) y ASE2 `F112…F129` (4 celdas) caen en un bloque detalle; `F353` cae en el detalle de ASE3.
3. Esas celdas son **slots de valor** en la v2 (sin `<f>`), y las anclas por empresa (`F66`, `F336`, `F529`…) también. La única corroboración por `<f>` disponible en la v2 es para 6 celdas: `F95`, `F86`, `F416`, `F417`, `F264` (consume `F112`) y `F415` (consume `F353`).
4. La cadena que citan los comentarios del mapa (p. ej. `F336 = F253+F237+F216−L237`) **no existe en ningún artefacto** (ni v1, ni v2, ni el manual de julio-2): proviene de la geometría legacy (`R-TRAMPA-ANCLA`) que T0-52 §8-b ya declaró.

**Consecuencia:** la pata "geometría del bloque oficial" de D2 **no es un tipador independiente** para estas celdas (condición **preexistente**, ya declarada en Plan 41 E1-2/E1-4 — "las anclas `REMUNERACION_*!D9:D13` quedan cortas/cero" — y en T0-52 §8-b/c). La tipabilidad del contrato se sostiene en:

- **(i)** la pata de sección **del lado de la fuente**: 180/180 filas tipadas, 0 empates, ≤3 secciones (§2);
- **(ii)** la pata de **fórmula consumidora** donde existe (6 celdas, §6.3) y, para el resto, la **semántica de operando documentada en el propio mapa** (`TotalMain`/`SubsMain` = bloque que cierra; `TotalMenor`/`SubsMenor` = primer bloque; `TotalMedio` = intermedio), que es la que las cadenas legacy describen;
- **(iii)** el **oráculo del manual**: 4/4 al centavo (§5.5) + identidad total en julio-1 (Q1 con conciliación) ⇒ **ninguna celda es ambigua**: si un rol estuviera mal, los Δ no coincidirían con §9 y el manual fallaría.

**No se dispara STOP:** ninguna celda quedó sin rol, ningún bloque presentó empate ni 4+ secciones, y el conjunto de Δ observado es **exactamente** el de la lista cerrada (§5.1+§5.3). La desviación es de *evidencia disponible en el artefacto*, no de tipabilidad, y se registra para el review de H2.

---

## 7. Veredicto H0

**VERDE — H1 habilitado**, con 2 puntos para el Ingeniero/review:

1. **Desviación de evidencia (D2.1)**: la pata de geometría oficial no existe para 34/56 celdas (slots de valor + direcciones fuera del bloque detalle del ASE). Se documenta con bytes y se sustituye por (i)+(ii)+(iii) del §6.
2. **Lista cerrada ampliada a 19 filas**: §9 (17 filas: 12 Δ + 5 identidades) + `F474`/`F477` de oct-1 ASE5 (Δ), normadas por D5 y autorizadas por el propio §9 ("se añade a la lista antes de H1 con evidencia").

**Fuera de alcance (declarado, no tocado):** las direcciones del mapa R1-empresa en la v2 (condición preexistente del frente E1-2), `R-SIGMA-Q2`, `R-MAPA-V2-R2-ASE4`, `ERR-SIN-SLOT` Q1 (Planes 41/46) y los 20 rojos preexistentes del reemplazo v2 `d55e621`.

---

## 8. Índice de evidencia

| Evidencia | Dónde |
|---|---|
| Matriz 52 grupos / 180 filas-empresa + partición/orden + tipabilidad | `SeccionPorOperandoR1T0Tests.S1_1_*` / `S1_6_*` |
| Tabla celda→sección (56) + coherencia rol↔operando | `SeccionPorOperandoR1T0Tests.S1_2_ContratoCeldaSeccion_CubreExactamenteElMapa` |
| Fórmulas consumidoras v2 + hojas inertes | `SeccionPorOperandoR1T0Tests.S1_2_EvidenciaV2_CeldasInertesYSusConsumidores` |
| Casos cardinales | `SeccionPorOperandoR1T0Tests.S1_3_CasosCardinales` |
| Inventario R2/R4 | `SeccionPorOperandoR1T0Tests.S1_4_InventarioR2R4` |
| Lista cerrada de Δ + identidades | `SeccionPorOperandoR1T0Tests.S1_5_ListaCerradaDelta_Exacta` |
| VisibleR1 por caso + identidad julio-1 | `SeccionPorOperandoR1T0Tests.S1_5_VisibleR1_PorCasoArbitrado` |
| Gate Σ (no empeora) | `SeccionPorOperandoR1T0Tests.S1_5_GateSigma_NoEmpeora` |
| Oráculo del manual | `SeccionPorOperandoR1T0Tests.S1_7_ManualArbitra_4Casos_AlCentavo` |
| Fuentes R1 (20 archivos) | `Docs/Prueba {Julio-1, Julio-2, Agosto-2, Octubre-1}/Insumos/*/Recaudoporcomponente_*.xlsx` (`Sheet1`) |
| Manual | `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx` · `Docs/Prueba Agosto-2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx` |
| Plantilla oficial v2 | `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx` (`Reporte Componentes R1`) |
| Diagnóstico T0-52 (base) | `Docs/T0-Plan52-Seccion-Operandos-Empresa-R1.md` |

**Nada de este documento modifica el repositorio**: el fixture es read-only y no escribe workbooks.
