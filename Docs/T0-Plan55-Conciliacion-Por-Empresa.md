# T0 — Plan 55: conciliación por empresa vs visible de bloque (cierre de bloque `2026101`)

| Campo | Valor |
|---|---|
| Plan | `plans/55 - conciliacion-por-empresa-v2-cierre-de-bloque.md` (**APROBADO 2026-10-23**) |
| Fase | **H0 (read-only, bloqueante)** — gate `H0 → H1` |
| Fecha | 2026-10-23 |
| Alcance ejecutado | T0.1 … T0.7. **Cero código de producción**: `Remuneracion.Core/` y `Remuneracion.Infrastructure/` no se tocan. Solo lectura BCL (ZIP+XML) de la plantilla v2/v1 y de las fuentes de octubre-1 de los ASE 1, 2 y 4 |
| Fixture | `Remuneracion.IntegrationTests/ConciliacionEmpresaCierreBloqueT0Tests.cs` — **81 casos, 81 verdes, 0 skip** |
| Invariantes | v2 `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` · v1 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` (**idénticos antes y después**) |
| Build | `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores** |
| Veredicto | **VERDE con 1 delta marcada PENDIENTE DE RATIFICACIÓN (D3, causa ii)** — H0 se detiene aquí; **H1 no se abre** sin el acto del Ingeniero |

> **Condición vinculante del §6 del plan:** la clasificación de una delta como causa **(ii) “la fuente no
> cuadra”** es un **veredicto de negocio**. Este documento **la propone con evidencia** y la deja marcada
> como *pendiente de ratificación*; no la declara final.
>
> **Regla de oráculo (R-ORACULO):** ninguna celda con `<f>` se usa por su `<v>` cacheado. Todo número de
> este documento viaja con su derivación (`fuente × <f>` v2 × mapa) — **CA0**.

---

## 0. Resumen ejecutivo

| # | ASE × reporte | Σ empresas | Visible de bloque | Δ | Causa propuesta | Estado |
|---|---|---:|---:|---:|---|---|
| **D1** | ASE1 · R1 | 14.815.023.989,82 | 14.815.021.974,75 | **2.015,07** | **(i)** dirección/sección | VERDE |
| **D2** | ASE4 · R1 | 10.474.277.163,42 | 10.473.967.813,42 | **309.350,00** | **(i)** sección Aplicación absorbida como “última” | VERDE |
| **D3** | ASE2 · R2 | 95.504.450,32 | 95.634.974,32 | **130.524,00** | **(ii)** etiqueta real sin dueño (`EAB`) | **PENDIENTE DE RATIFICACIÓN** |
| **D4** | ASE1 · R4 | −134.910,00 | −3.833.780,00 | **3.698.870,00** | **(i)** mapa sin entrada (`ENEL`) | VERDE (encuadre a ratificar, §5.4) |
| **D5** | ASE2 · R4 | −20.858.585,45 | −21.803.648,45 | **945.063,00** | **(i)** mapa sin entrada (`OCCIDENTE`) | VERDE (encuadre a ratificar, §5.4) |

Las cinco deltas se **reproducen al peso por bytes** desde las fuentes de octubre-1 (§4) con la misma
semántica del lector de producción, y coinciden con los cinco ítems `ERR-VALIDACION` de la corrida real
(`plans/55` §0.2). **Ninguna** requiere ajustar valores, tolerancias ni mapas.

**Hallazgo estructural que ordena el resto:** las 112 celdas del mapa por empresa **no alimentan el
visible de bloque de su propio ASE** (§3.3) y, en la corrida 5-ASE, las 56 celdas R1 **ni siquiera se
escriben** (`OpenXmlPlantillaWriter.cs:266` `espejoR1 = bloquesEspejo.Count == 5` → `:323`
`EscribirCeldasEmpresa(..., omitirR1: true)` → `:1250` `if (!omitirR1 && …)`). Es decir: **las deltas R1
no son un defecto de dirección de escritura, son un defecto de la descomposición en dominio** (qué
sección aporta al visible por empresa). Se documenta con bytes; **no se corrige en H0**.

---

## 1. Método (reproducible)

1. **BCL puro**: `System.IO.Compression.ZipArchive` + `System.Xml.Linq`; `workbook.xml` + rels para
   resolver la hoja por **nombre**. Sin Excel/COM, sin OpenXml SDK, sin ExcelDataReader.
2. **Rutas explícitas** (prohibido `Insumos.Raiz()` en lo nuevo): el fixture resuelve la raíz por su
   cuenta (`AGENTS.md`) y declara carpeta + nombre de archivo completo por ASE.
3. **Naturaleza de celda**: presencia del elemento `<f>` (con su texto; `t="shared"` se anota como
   fórmula compartida **sin texto propio**), `<v>` solo como valor de literales.
4. **Espejo R1**: misma ventana que `ConstruirSecuenciaEspejoR1` (primera fila `E='Vlr Servicio'` …
   fila invariante `A='Total'` con `B` vacío, sin filas fantasma).
5. **Tipado de sección**: predicado vigente `ContratoSeccionPeriodFreeR1.SeccionPorCierre` (Plan 46) y
   resolución de operando con `ResolvedorOperandosEmpresaR1` (Plan 53) — **la misma semántica de
   producción, no una reimplementación**.
6. **Cero escrituras**: no se abrió ninguna plantilla en modo escritura ni se generaron salidas.

---

## 2. G-SHA / G-FORMULA (T0.7)

| Plantilla | SHA-256 antes | SHA-256 después | Veredicto |
|---|---|---|---|
| v2 `Formato Remuneracion 2026xx-x Total vf.xlsx` | `0F192AD2…7F2DAB` | `0F192AD2…7F2DAB` | **IDÉNTICO** |
| v1 `Formato Remuneracion 2026xx-x Total.xlsx` | `3086AC31…C5214` | `3086AC31…C5214` | **IDÉNTICO** |

Las **10 `<f>` de R1** siguen byte-idénticas (`T0_7_FormulasR1_ByteIdenticas`, 10/10):
`F75`, `F77` (ASE1) · `F254`, `F256` (ASE2) · `F415`, `F417` (ASE3) · `F568`, `F570` (ASE4) ·
`F680`, `F682` (ASE5).

**Cero `InsertRow`/`DeleteRow`, cero `<f>` nueva.** H0 no escribe archivos.

---

## 3. T0.1 — destinos por bytes (S1.1 / CA4)

### 3.1 Geometría del template (medida, no heredada)

`Reporte Componentes R1` (`xl/worksheets/sheet10.xml`, rId10) tiene un bloque **detalle** y un bloque
**resumen** por ASE. Cabeceras y cierres medidos (`B` = nombre del ASE, `A='Total'` = cierre):

| ASE | Cabecera | Detalle | Cierre | Resumen |
|---|---|---:|---:|---|
| 1 PROMOAMBIENTAL | `B4` | 5–72 | `A72` | 75–106 |
| 2 LIME | `B108` | 109–210 | `A210` | 211–285 |
| 3 CIUDAD LIMPIA | `B287` | 288–355 | `A355` | 358–449 |
| 4 BOGOTA LIMPIA | `B448` | 449–550 | `A550` | 551–… |
| 5 AREA LIMPIA | `B601` | 602–676 | `A676` | 678–… |

R2 (`sheet11.xml`): bloques por `B` = `B1` / `B65` / `B173` / `B291` / `B402`.
R4 (`sheet12.xml`): bloques por `B` = `B1` / `B91` / `B192` / `B229` / `B344`.

### 3.2 Destino de cada operando por empresa y su naturaleza

`<f>` = fórmula (texto entre comillas), `L` = literal sin fórmula. Todos los `L` valen `0` salvo donde se
indica `∅` (celda presente sin `<v>`).

**R1 — `Reporte Componentes R1` (56 celdas)**

| ASE | Empresa | Celda (operando) | Naturaleza |
|---|---|---|---|
| 1 | 2 ENEL | `F33` (SubsMain), `F14` (TotalMain), `L14` (EspecialesMain) | 3× L |
| 1 | 4 OCCIDENTE | `F40` (SubsMain), `F24` (TotalMain), `L24` (EspecialesMain), `F29` (SubsMenor), `F9` (TotalMenor), `L9` (EspecialesMenor) | 6× L |
| 2 | 1 RECIP (`NUEVO ESQUEMA`) | `F102` (TotalMain), `F122` (SubsMain), **`F86`** (Block5Total) | 2× L · **1× `<f>F62`** |
| 2 | 2 ENEL | `F117` (SubsMain), **`F95`** (TotalMain), `F83` (TotalMenor), **`L95`** (EspecialesMain) | 2× L · **1× `<f>F69+F46+F15-L15-L46`** · **1× `<f>` compartida (`t="shared" si=4`)** |
| 2 | 4 OCCIDENTE | `F129` (SubsMain), `F112` (TotalMain), `F89` (TotalMenor, ∅), `L112` (EspecialesMain) | 4× L (∅ en `F89`) |
| 3 | 2 ENEL | `F246`, `F227`, `F213`, `L227` | 4× L (∅ en todas) |
| 3 | 4 OCCIDENTE | `F253`, `F237`, `F216`, `L237`, `F242`, `F222` | 6× L (∅ en todas) |
| 4 | 1 RECIP (`NUEVO ESQUEMA`) | `F409` (SubsMain), `F380` (TotalMain), `L380` (EspecialesMain) | 3× L (∅ en todas) |
| 4 | 2 ENEL | `F404`, `F373`, `F353`, `L373`, `F395`, `F362` | 6× L (∅ salvo `F353`) |
| 4 | 4 OCCIDENTE | **`F416`** (SubsMain), `F390` (TotalMain), `F356` (TotalMenor), `L390` (EspecialesMain), `F400` (SubsMenor), `F368` (TotalMedio) | 5× L (∅ en todas) · **1× `<f>F353`** |
| 5 | 2 ENEL | `F502`, `F483`, `F474`, `L483` | 4× L |
| 5 | 3 ENERBIT | `F505`, `F487`, `L487` | 3× L |
| 5 | 4 OCCIDENTE | `F512`, `F497`, `F477`, `L497` | 4× L |

**R2 — `Rem. Anticipos R2` (36 celdas)**

| ASE | Empresa | Celda (operando) | Naturaleza |
|---|---|---|---|
| 1 | 2 ENEL | `E24` (SubsCont), `E10` (Total), `K10` (Especiales) | 3× L |
| 1 | 4 OCCIDENTE | `E20` (Total), `E31` (SubsCont), `K20` (Especiales) | 3× L |
| 2 | 1 RECIP | `E104` (SubsCont), `E83` (Total), `K83` (Especiales) | 3× L |
| 2 | 2 ENEL | `E97`, `E74`, `K74` | 3× L |
| 2 | 4 OCCIDENTE | `E93`, `E111`, `K93` | 3× L |
| 3 | 2 ENEL | `E185`, `E171`, `K171` | 3× L (∅ en `K171`) |
| 3 | 4 OCCIDENTE | `E181`, `E192`, `K181` | 3× L |
| 4 | 1 RECIP | `E319` (SubsCont), **`E285`** (Total), **`K285`** (Especiales) | 1× L · **1× `<f>E192+E203-K192`** · **1× `<f>` compartida (`si=9`)** |
| 4 | 2 OCCIDENTE | `E312`, `E289`, `K289` | 3× L (∅ en `K289`) |
| 4 | 4 OCCIDENTE | `E308`, `E326`, `K308` | 3× L |
| 5 | 2 ENEL | `E407`, `E389`, `K389` | 3× L (∅ en `E389`, `K389`) |
| 5 | 4 OCCIDENTE | `E403` (¡texto de cabecera `Total`!), `E417`, `K403` (¡texto `Especiales`!) | 3× L de texto |

**R4 — `Reversion Pagos R4` (20 celdas)**

| ASE | Empresa | Celda | Naturaleza |
|---|---|---|---|
| 1 | 4 OCCIDENTE | `D14`, `P14` | 2× L |
| 2 | 1 RECIP | `D104`, `P104` | 2× L |
| 2 | 2 ENEL | `D98`, `P98` | 2× L |
| 3 | 2 ENEL | `D199`, `P199` | 2× L |
| 3 | 4 OCCIDENTE | `D205`, `P205` | 2× L |
| 4 | 1 RECIP | `D242`, `P242` | 2× L |
| 4 | 2 ENEL | `D236`, `P236` | 2× L |
| 4 | 4 OCCIDENTE | `D248`, `P248` | 2× L |
| 5 | 2 ENEL | `D351`, `P351` | 2× L |
| 5 | 4 OCCIDENTE | `D357`, `P357` | 2× L |

### 3.3 Lista cerrada: celdas destino que son FÓRMULA en la v2 (clase `R-MAPA-V2-R2-ASE4`)

| Celda | Texto `<f>` | Mapa | Efecto si se escribe |
|---|---|---|---|
| `R1!F86` | `F62` | (RECIP, ASE2) Block5Total | `ERR-PLANTILLA` (guard) |
| `R1!F95` | `F69+F46+F15-L15-L46` | (ENEL, ASE2) TotalMain | `ERR-PLANTILLA` |
| `R1!L95` | — (compartida `si=4`) | (ENEL, ASE2) EspecialesMain | `ERR-PLANTILLA` |
| `R1!F416` | `F353` | (OCCIDENTE, ASE4) SubsMain | `ERR-PLANTILLA` |
| `R2!E285` | `E192+E203-K192` | (RECIP, ASE4) Total | `ERR-PLANTILLA` |
| `R2!K285` | — (compartida `si=9`) | (RECIP, ASE4) Especiales | `ERR-PLANTILLA` |

**Ninguna** celda destino es fórmula en R4 (20/20 literales).

> **Hallazgo T0 (nuevo, medido aquí):** de las 112 celdas destino del mapa por empresa, **ninguna es
> término de la `<f>` del visible de bloque de su propio ASE** (teorema asertado por ASE × reporte en
> `T0_1_NingunDestinoEsTerminoDelVisibleDeBloque`, 15/15 verde). Además, **34/56 celdas R1 caen en el
> bloque de OTRO ASE** (ya declarado en `Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md` §6) y una
> de ellas produce **contaminación cruzada**: `F353` es destino de `(ENEL, ASE4)` y a la vez término del
> TOT_OPT de **ASE3** (`F415 = F330+F353+F299-L299-L330`).

---

## 4. T0.2 — censo de octubre-1 y derivación de los 5 visibles de bloque

Fuentes (rutas explícitas del fixture, hoja `Sheet1`):

| ASE | Carpeta | R1 | R2 | R4 |
|---|---|---|---|---|
| 1 | `1-Promoambiental` | `Recaudoporcomponente_…_20261016174133882.xlsx` | `RerpoteDetalleSaldosaFavor_…_20261016174159251.xlsx` | `ReversiónPorComponente_…_20261016174217316.xlsx` |
| 2 | `2-Lime` | `Recaudoporcomponente_…_20261016174350151.xlsx` | `RerpoteDetalleSaldosaFavor_…_2026101617441360.xlsx` | `ReversiónPorComponente_…_20261016174430392.xlsx` |
| 4 | `4-Bogota Limpia` | `Recaudoporcomponente_…_20261016174948502.xlsx` | `RerpoteDetalleSaldosaFavor_…_202610161750610.xlsx` | `ReversiónPorComponente_…_20261016175022764.xlsx` |

### 4.1 Censo R1

| ASE | filas `Mes/Total` (`B=Mes ∧ C=Total`) | filas `Aplicacion/Total` | cierre (`A=Total`, `B` vacío) |
|---|---|---|---|
| 1 | **8** (`F=−2.015,07` · `L=0`), **33** (`F=18.025.191.391,39` · `L=19.931.693,03`), **52** (`F=−3.190.235.708,54` · `L=0`) | 18 (`636.047,07`), 41 (`12.822,93`) | 54 (`14.835.602.537,78`) |
| 4 | **11** (`F=−30.180,94` · `L=0`), **46** (`F=10.776.632.651,21` · `L=8.643.135,58`), **70** (`F=−293.991.521,27` · `L=0`) | 31 (`23.031.582,34`), 59 (`−1.474.405,34`) | 72 (`10.504.168.126,00`) |

Identidad verificada en bytes: la fila `A='Total'` de la fuente **no es un visible** — es
`Σ Mes + Σ Aplic` y **no resta** `SERVICIO ESPECIALES` (coincide con T0-Plan54 §2).

### 4.2 Censo R2 / R4

| ASE | reporte | etiquetas de cierre | fila `Total` |
|---|---|---|---|
| 2 | R2 | `Componente/Total` fila 33 (`E=116.410.420,67`, `K=70.736,95`) · `Subs/Cont/Total` fila 56 (`E=−20.704.709,40`) | 57 |
| 1 | R4 | `ENEL` fila 10 (`D=−3.698.870,00`) · `OCCIDENTE` fila 16 (`D=−134.910,00`) | 17 (`−3.833.780,00`) |
| 2 | R4 | `ENEL` fila 10 (`D=−20.858.585,45`) · `OCCIDENTE` fila 16 (`D=−945.063,00`) | 17 (`−21.803.648,45`) |

### 4.3 Los 5 visibles de bloque (derivación, CA0)

| ASE × reporte | Derivación de bytes | Visible | `<f>` v2 del bloque |
|---|---|---:|---|
| ASE1 · R1 | `ΣF(Mes) − ΣL(Mes sin la última)` = `14.834.953.667,78 − 19.931.693,03` | **14.815.021.974,75** | `F75 = F47+F70+F16-L16-L47` |
| ASE4 · R1 | `10.482.610.949,00 − 8.643.135,58` | **10.473.967.813,42** | `F568 = F511+F548+F460-L460-L511` |
| ASE2 · R2 | `E(Componente/Total) + E(Subs/Cont/Total) − K(Componente/Total)` = `116.410.420,67 − 20.704.709,40 − 70.736,95` | **95.634.974,32** | `E155 = E101+E123-K101` |
| ASE1 · R4 | `D(fila 17)` | **−3.833.780,00** | `D73 = D15-P15` |
| ASE2 · R4 | `D(fila 17)` | **−21.803.648,45** | `D174 = D111-P111` |

---

## 5. T0.3 — tabla de atribución (una fila por delta)

> **Convención:** “Σ empresas” es el agregado por empresa **del dominio** (`LeerConciliacionEmpresas` +
> `CalcularVisibleR*`), reconstruido aquí por bytes; “visible de bloque” es el valor del dominio que
> `ValidarSigmaEmpresas` compara (`WorkbookLeafCoherence.cs:105-120`), reconstruido aquí por bytes.
> Las citas de template son `hoja!celda` de la **v2 vigente** (`0F192AD2…`).

| Δ | ASE | Componente | Δ medida | Empresa(s) implicada(s) | Cita de la FUENTE (bytes) | Cita del TEMPLATE v2 | Naturaleza de la celda destino | Causa propuesta | Evidencia de bytes que la sostiene |
|---|---|---|---:|---|---|---|---|---|---|
| **D1** | ASE1 | R1 | **2.015,07** | ENEL (2) — la 1.ª sección es toda de ENEL; OCCIDENTE (4) no aporta a la sección | `Recaudoporcomponente…4133882` · `Sheet1!C7:F7` → `C7=ENEL D7=Total F7=−2.015,07` y `Sheet1!B8:F8` → `B8=Mes C8=Total F8=−2.015,07` (`L7=L8=0`) | `Reporte Componentes R1!F16` (literal `0`, slot `Mes0`), `!L16` (`0`), `!F75 = F47+F70+F16-L16-L47`; destinos del par: `!F33`, `!F14`, `!L14` (literales, fuera de `F75`) | literal (destino) / `<f>` (consumidor) | **(i)** la 1.ª sección (`Oportuno#1`, cerrada por `Mes/Total` fila 8) **entra en el bloque** (término `F16`) y **no entra en el agregado por empresa**: la cadena de ASE1 es de 3 términos y el mapa `(2,1)` no declara `TotalMenor` | Bloque `14.834.953.667,78 − 19.931.693,03 = 14.815.021.974,75`; Σ empresas `ENEL 14.471.401.705,97 + OCC 343.622.283,85 = 14.815.023.989,82`; Δ `= +2.015,07` = \|fila 8\|. Σ de filas-etiqueta de las 3 secciones `Oportuno` **= bloque al peso** ⇒ la fuente cuadra y el que no cuadra es el mapa |
| **D2** | ASE4 | R1 | **309.350,00** | ReciprocidadEaab (1) — label de fuente `NUEVO ESQUEMA` | `Recaudoporcomponente…174948502` · `Sheet1!C22:F22` → `C22=NUEVO ESQUEMA D22=Total F22=356.584,67`; `Sheet1!C53:F53` → `F53=−47.234,67`; `Sheet1!B31:F31` → `Aplicacion/Total F31=23.031.582,34`; filas `Mes/Total` 11/46/70 | `Reporte Componentes R1!F380`, `!F409`, `!L380` (literales, filas 380/409 dentro del bloque/resumen de **ASE3** 358–449); visible del ASE4 `!F568 = F511+F548+F460-L460-L511` | literal | **(i)** el operando resuelve `UltimaOportuna` = **la única sección presente**, que es `Aplicacion#1`; el TOT_OPT del bloque suma **solo filas `Mes/Total`** y no recorre Aplicación ⇒ el dinero entra en Σ empresas y no en el bloque | Σ `9.882.391.541,17 + 591.576.272,25 + 309.350,00 = 10.474.277.163,42`; bloque `10.482.610.949,00 − 8.643.135,58 = 10.473.967.813,42`; Δ `= 356.584,67 − 47.234,67 = 309.350,00`. El label **no tiene ninguna fila en secciones `Oportuno`** (asertado) ⇒ no es huérfano ni duplicado |
| **D3** | ASE2 | R2 | **130.524,00** | Sin dueño en el catálogo: etiqueta `EAB` (empresa 5 `EaabCiudadLimpia` espera `CIUDAD LIMPIA-ACUEDUCTO` y no aparece en la fuente) | `RerpoteDetalleSaldosaFavor…17441360` · `Sheet1!B7:E7` → `B7=EAB C7=Total E7=149.137,87` (Componente) y `Sheet1!B36:E36` → `E36=−18.613,87` (Subs/Cont) | `Rem. Anticipos R2!B70` = `EAB`, `!C70` = `Total`, `!E70`/`!K70` literales `0` (slot libre, **sin entrada en el mapa**); bloque `!E155 = E101+E123-K101` | literal | **(ii) la fuente no cuadra** — etiqueta **real** sin dueño en el vocabulario de fuentes (`R-ETIQUETAS-SIN-DESTINO`, T0-Plan50 §3.2/§4.2). **PENDIENTE DE RATIFICACIÓN DEL INGENIERO**; prohibido inventar celda destino en este plan | Bloque `116.410.420,67 − 20.704.709,40 − 70.736,95 = 95.634.974,32`; Σ `11.631.857,27 + 39.342.901,74 + 0 + 44.529.691,31 + 0 = 95.504.450,32`; Δ `= 149.137,87 − 18.613,87 = 130.524,00` **exacto**. `EditablesR2PorEmpresa` no tiene clave `(5,2)` ni ninguna etiqueta `EAB` |
| **D4** | ASE1 | R4 | **3.698.870,00** | Enel (2) | `ReversiónPorComponente…174217316` · `Sheet1!A10:D10` → `A10=ENEL B10=Total D10=−3.698.870,00`; `Sheet1!A16:D16` → `OCCIDENTE D16=−134.910,00`; `Sheet1!A17:D17` → `Total D17=−3.833.780,00` | `Reversion Pagos R4!A8` = `ENEL`, `!B8` = `Total`, `!D8`/`!P8` literales `0` (**slot existente**), resumen `!D79 = D8-P8`, visible del bloque `!D73 = D15-P15`; el mapa **no** declara `(2,1)` | literal | **(i)** cobertura del mapa: falta la entrada de `ENEL` en `EditablesR4PorEmpresa` para ASE1. La fuente cuadra al peso y **el template sí tiene el slot** | Bloque `−3.833.780,00`; Σ `= −134.910,00` (solo OCCIDENTE); Δ `= 3.698.870,00` = fila 10. Identidad de la fuente: `−3.698.870,00 + (−134.910,00) = −3.833.780,00` |
| **D5** | ASE2 | R4 | **945.063,00** | OccidenteDirecta (4) | `ReversiónPorComponente…174430392` · `Sheet1!A10:D10` → `ENEL D10=−20.858.585,45`; `Sheet1!A16:D16` → `OCCIDENTE D16=−945.063,00`; `Sheet1!A17:D17` → `Total D17=−21.803.648,45` | `Reversion Pagos R4!A110` = `OCCIDENTE`, `!B110` = `Total`, `!D110`/`!P110` literales `0` (**slot existente**), resumen `!D186 = D110-P110`, visible del bloque `!D174 = D111-P111`; el mapa **no** declara `(4,2)` | literal | **(i)** cobertura del mapa: falta la entrada de `OCCIDENTE` en `EditablesR4PorEmpresa` para ASE2. La fuente cuadra al peso y el template tiene el slot | Bloque `−21.803.648,45`; Σ `= −20.858.585,45` (`RECIP` está en el mapa `(1,2)` pero **no tiene fila** en la fuente ⇒ 0); Δ `= 945.063,00` = fila 16. Identidad: `−20.858.585,45 + (−945.063,00) = −21.803.648,45` |

### 5.1 Firma de cada causa (contraste con la taxonomía cerrada del §D2 del plan)

| Δ | Firma del plan | ¿Se cumple? |
|---|---|---|
| D1 | (i) “`Σ etiquetas-fuente == Total-fuente` (la fuente cuadra) pero `Σ empresas ≠ bloque` (el mapa no)” | **Sí** (§5 fila D1) |
| D2 | (i) “sección esperada del operando ≠ sección que la `<f>` suma” | **Sí**: `UltimaOportuna` → `Aplicacion#1`, y el TOT_OPT suma `Mes` |
| D3 | (ii) “`Σ etiquetas-reconocidas ≠ Total-fuente` por el dinero exacto de la etiqueta sin dueño” | **Sí**: 130.524,00 = `EAB` |
| D4 | (i) “duplicado/huérfano”, `Σ etiquetas == Total` | **Sí** (§5 fila D4) |
| D5 | (i) idem | **Sí** (§5 fila D5) |

### 5.2 Lo que NO es causa de ninguna delta (descartado por bytes)

- **No** es un problema de dirección de escritura: las 56 celdas R1 **no se escriben** en la corrida
  5-ASE (`OpenXmlPlantillaWriter.cs:266/323/1250`) y, aunque se escribieran, **ninguna** es término del
  TOT_OPT de su ASE (§3.3). Las deltas R1 son de **descomposición en dominio**.
- **No** es la tolerancia (±0,5): las 5 deltas la superan por 3–7 órdenes de magnitud.
- **No** es el Plan 54: la semántica de agregado y los mapas R2/R4 quedaron intactos (Plan 54 §8 G-PROC).

### 5.3 Las cinco Σ de empresas (derivación al peso)

| ASE × reporte | Empresa | Operandos (fuente → celda del mapa) | Visible por empresa |
|---|---|---|---:|
| ASE1 · R1 | 2 ENEL | `F14` Ultima `17.734.947.270,61` + `F33` Ultima `−3.253.607.640,54` − `L14` `9.937.924,10` | 14.471.401.705,97 |
| ASE1 · R1 | 4 OCCIDENTE | `F24` `290.244.120,78` + `F40` `63.371.932,00` − `L24` `9.993.768,93` (`F9/F29/L9` **no** entran: cadena de 3 términos) | 343.622.283,85 |
| **ASE1 · R1** | | **Σ** | **14.815.023.989,82** |
| ASE4 · R1 | 1 RECIP | `F380` `356.584,67` + `F409` `−47.234,67` − `L380` `0` | 309.350,00 |
| ASE4 · R1 | 2 ENEL | `F373` `10.331.368.139,96` + `F404` `−440.994.165,67` + `F353` `−10.060,29` − `L373` `7.972.372,83` | 9.882.391.541,17 |
| ASE4 · R1 | 4 OCCIDENTE | `F390` `445.264.511,25` + `F416` `147.002.644,40` + `F356` `−20.120,65` − `L390` `670.762,75` | 591.576.272,25 |
| **ASE4 · R1** | | **Σ** | **10.474.277.163,42** |
| ASE2 · R2 | 1 RECIP | `E83` `16.015.076,93` + `E104` `−4.383.219,66` − `K83` `0` | 11.631.857,27 |
| ASE2 · R2 | 2 ENEL | `E74` `58.544.042,33` + `E97` `−19.130.403,64` − `K74` `70.736,95` | 39.342.901,74 |
| ASE2 · R2 | 3 ENERBIT | sin filas del label ⇒ 0 | 0,00 |
| ASE2 · R2 | 4 OCCIDENTE | `E93` `41.702.163,54` + `E111` `2.827.527,77` − `K93` `0` | 44.529.691,31 |
| ASE2 · R2 | 5 EAAB-CL | sin filas (`CIUDAD LIMPIA-ACUEDUCTO` no aparece) ⇒ 0 | 0,00 |
| **ASE2 · R2** | | **Σ** | **95.504.450,32** |
| ASE1 · R4 | 4 OCCIDENTE | `D14` `−134.910,00` − `P14` `0` | −134.910,00 |
| **ASE1 · R4** | | **Σ** | **−134.910,00** |
| ASE2 · R4 | 1 RECIP | sin fila en la fuente ⇒ 0 | 0,00 |
| ASE2 · R4 | 2 ENEL | `D98` `−20.858.585,45` − `P98` `0` | −20.858.585,45 |
| **ASE2 · R4** | | **Σ** | **−20.858.585,45** |

### 5.4 Nota de encuadre para D4/D5 (declarada, no decidida)

`plans/55` §7 lista bajo `R-ETIQUETAS-SIN-DESTINO` el caso “`ENEL` sin destino en R4 de ASE1/3/5”, y
T0-Plan50 §4.2 anticipó “ASE1 R4 (→ 3.698.870 = `ENEL` sin destino)”. **El hecho es el mismo** que aquí se
clasifica como **(i)**; lo que difiere es el nombre del cajón. El criterio operativo del §D2 del plan
manda por la **firma**: en D4/D5 la fuente **cuadra al peso** (`Σ etiquetas-fuente == Total-fuente`), el
label **sí** pertenece al catálogo y **el template tiene el slot** (`R4!D8`, `R4!D110`) ⇒ la reparación es
técnica (cobertura del mapa), no un veredicto de negocio. **Si el Ingeniero prefiere encuadrarlas como
(ii), el veredicto de negocio no cambia la reparación** — pero sí exige su ratificación expresa. Queda
declarado para el acto de aprobación; **no bloquea H1 por sí mismo**.

---

## 6. T0.4 — los dos delatores

### 6.1 H-ASE4-R1 (`309.350,00` = `ReciprocidadEaab`): **CONFIRMADA la identidad, REFUTADA la lectura “huérfano/duplicado”**

- **Identidad numérica:** `356.584,67` (componente, `Sheet1!F22`) `+ (−47.234,67)` (subsidio,
  `Sheet1!F53`) `= 309.350,00`. Exacto, sin tolerancia.
- **No es huérfano ni duplicado:** el dinero existe en la fuente, una sola vez, en la sección
  `Aplicacion#1` (`Sheet1!B31` `F=23.031.582,34` cierra las filas 13–30). `NUEVO ESQUEMA` **no tiene
  ninguna fila en las secciones `Oportuno`** de ASE4 (filas `Mes/Total` 11, 46, 70): solo aparece en 22
  (componente) y 53 (subsidio).
- **El bloque no lo recorre:** el TOT_OPT del bloque suma únicamente las filas `Mes/Total`
  (`F568 = F511+F548+F460-L460-L511`), es decir las tres secciones `Oportuno`. La sección `Aplicacion`
  no participa.
- **El operando lo absorbe:** `ResolvedorOperandosEmpresaR1` con `UltimaOportuna` devuelve “la última
  sección presente del bloque **de la empresa**”, que para RECIP es la única (`Aplicacion#1`) ⇒ el
  operando por empresa sí lo suma.
- **Consecuencia:** causa **(i)**; el destino (`F380/F409/L380`) es literal pero vive en el bloque/resumen
  de ASE3 (358–449) y no es término de `F568`.

### 6.2 H-ASE1-R1 (`2.015,07` = |primera fila `Mes`|): **CONFIRMADA la identidad, REFUTADA la dirección**

- **Identidad numérica:** |`Sheet1!F8`| = |−2.015,07| = Δ. Exacto.
- **La hipótesis enunciaba** “la sección *primera* entra en el agregado por empresa y **no** en el
  bloque, o al revés”. Los bytes dicen **al revés de la primera alternativa**: la sección *primera*
  (`Oportuno#1`, cerrada por la fila `Mes/Total` **8**) **SÍ entra en el bloque** — es el término `F16`
  de `F75 = F47+F70+F16-L16-L47` — y **NO entra en el agregado por empresa**.
- **Por qué no entra:** (a) la cadena T0 de ASE1 en `CalcularVisibleR1` es de **3 términos**
  (`totalMain + subsMain − especialesMain`) e **ignora `totalMenor`**; (b) el mapa `(2,1)` (ENEL) **no
  declara `TotalMenor`**. Resultado: de las 3 secciones de ENEL en Totales
  (`Oportuno#1 = −2.015,07`, `Aplicacion#1 = 546.059,64`, `Oportuno#2 = 17.734.947.270,61`) el agregado
  solo toma la última.
- **La fuente cuadra:** la suma de filas-etiqueta de las tres secciones `Oportuno`
  (`−2.015,07 + 18.025.191.391,39 − 3.190.235.708,54`) es **exactamente** el visible de bloque.
- **Consecuencia:** causa **(i)** de sección (no (ii)).

---

## 7. T0.5 — veredicto `R-RECORTE-Q2` (T4 / D6): **NO se levanta**

- `R2!E285` = `<f>E192+E203-K192</f>` y `R2!K285` = `<f t="shared" si="9"/>` en la **v2 vigente**: el
  destino del par `(RECIP, ASE4)` de R2 es **fórmula** ⇒ escribirlo aborta con `ERR-PLANTILLA`
  (`OpenXmlPlantillaWriter.cs:1993-2005`, `EscribirValorNumerico` rechaza `cell.CellFormula != null`).
- Además hay **4 destinos R1 fórmula** (parte del mismo barrido): `F86`, `F95`, `L95`, `F416`.
- **El T0 no respalda el levantamiento** ⇒ `R-RECORTE-Q2` **sigue vivo** con su gate (`D6`), y
  `ProcesadorPeriodo.cs:198-217` **no se toca**. `R-SIGMA-Q2` (T0-Plan50 §4.2) sigue vigente: el gate
  Sigma **no** se reactiva en Q2 aunque se levante el recorte.

### 7.1 Bloqueador siguiente (declarado, no implementado)

Aun si H1 cerrara las 5 deltas, la corrida 5-ASE de `2026101` **seguiría en rojo**: `EscribirCeldasEmpresa`
escribe **R2/R4 por empresa siempre** (no los omite con `espejoR1`) y `(1,4)` cae en `E285`/`K285`
(fórmulas) ⇒ `ERR-PLANTILLA`. **H1 debe re-congelar también esos dos destinos R2.** Queda como
dependencia explícita de H1 (no es una delta del gate: es el guard de plantilla).

---

## 8. T0.6 — cobertura del gate (F5) y lista cerrada de Δ de suite

| Verificación | Resultado |
|---|---|
| `ValidarSigmaEmpresas` en `Remuneracion.IntegrationTests/TestResults/p53-h2-suite.trx` | **0 ocurrencias** |
| mensaje “visibles de empresas” en el mismo trx | **0 ocurrencias** |
| contadores de la línea base | `total="1289"` · `passed="1222"` · `failed="67"` · `skipped=0` |
| `p54-H3-suite.trx` | `total="1374"` · `passed="1307"` · `failed="67"` |

**Lista cerrada de Δ de la suite en H0:** **ninguna baja, ningún rojo nuevo.** El único Δ es el **alta**
del fixture de este T0 (`ConciliacionEmpresaCierreBloqueT0Tests`: 81 casos, 81 verdes, 0 skip). Cualquier
otro movimiento de la suite ⇒ **STOP**.

Corrida completa de H0 (`TestResults/p55-H0-suite.trx`):

| Métrica | Valor medido | Línea base `p54-H3-suite.trx` |
|---|---:|---:|
| total | **1455** | 1374 |
| passed | **1388** | 1307 |
| failed | **67** | 67 |
| skipped | **0** | 0 |
| altas | **+81** (todas del fixture T0, todas verdes) | — |
| bajas | **0** | — |
| set-diff de rojos | **IDÉNTICO** (67 = 67, mismos nombres) | — |

---

## 9. Índice de evidencia (trazabilidad test ↔ afirmación)

| Afirmación | Test |
|---|---|
| SHAs v1/v2 intactos | `T0_7_PlantillasOficiales_ShaIntacto` |
| Las 10 `<f>` de R1 byte-idénticas | `T0_7_FormulasR1_ByteIdenticas` (10) |
| Destinos 112 celdas + naturaleza + simetría con el mapa | `T0_1_DestinosPorBytes_CoincidenConElMapaYSuNaturaleza` |
| Lista cerrada de 6 destinos fórmula | `T0_1_DestinosFormula_ListaCerradaSeis` |
| Ningún destino es término del visible de su ASE | `T0_1_NingunDestinoEsTerminoDelVisibleDeBloque` (15) |
| Contaminación cruzada `F353` | `T0_1_ContaminacionCruzada_F353_EsDestinoDeAse4YTerminoDeAse3` |
| Censo octubre-1 de los 5 bloques | `T0_2_CensoOctubre1_PorBytes` (5) |
| Operandos por empresa (derivación fuente × mapa) | `T0_3_OperandoPorEmpresa_DerivacionFuenteMapa` (38) |
| Las 5 deltas al peso | `T0_3_DeltaMedida_SeReproducePorBytes` (5) |
| H-ASE4-R1 | `T0_4_Hipotesis_Ase4R1_ReciprocidadEsAplicacionYNoOportuno` |
| H-ASE1-R1 | `T0_4_Hipotesis_Ase1R1_PrimeraSeccionEntraEnBloqueYNoEnEmpresa` |
| `R-RECORTE-Q2` no se levanta | `T0_5_RecorteQ2_NoSeLevanta_DestinoE285EsFormula` |
| Gate sin cobertura (F5) | `T0_6_GateSigmaEmpresas_SinCoberturaEnLaLineaBase` |

---

## 10. Puerta H0 → H1

| Criterio | Estado |
|---|---|
| Fixture T0 verde (81/81, 0 skip) | ✅ |
| S1.1 destinos v2 por bytes + CA4 (lista cerrada de fórmulas) | ✅ |
| S1.2 censo octubre-1 congelado | ✅ |
| S1.3 clasificación D1-D5 con derivación completa | ✅ (D3 = **(ii) pendiente de ratificación**) |
| S1.4 H-ASE4-R1 y H-ASE1-R1 confirmadas/refutadas por bytes | ✅ |
| S1.5 veredicto `R-RECORTE-Q2` | ✅ (vivo, no se levanta) |
| S1.6 lista cerrada de Δ de suite | ✅ (solo el alta del fixture) |
| G-SHA / G-FORMULA | ✅ |
| G-H0-READONLY: diff vacío en `Remuneracion.Core/` + `Remuneracion.Infrastructure/` | ✅ (cero código de producción) |
| Build 0 warnings | ✅ |

**Veredicto: H0 VERDE con una delta pendiente de ratificación (D3).** Por la condición vinculante del §6
del plan, **H0 se detiene aquí**: la clasificación **(ii)** de D3 y la apertura de **H1** requieren el
acto expreso del Ingeniero. Nada de este documento modifica el repositorio más allá del fixture read-only
y de este documento.
