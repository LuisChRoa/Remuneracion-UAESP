---
titulo: Plan 56 - Jerarquia seccion-subseccion de la plantilla R1 (identidad de cuadro A/B)
estado: APROBADO 2026-10-23 (gate §6 ratificado; H0 habilitado, H1 condicionada al T0)
fecha: 2026-10-23
numero: 56
planes_ocupados: 01..55 ocupados; este es el 56
padre: plans/55 - conciliacion-por-empresa-v2-cierre-de-bloque.md
madre_doctrina: .opencode/project-context.md §"Modelo de la plantilla estándar" → "Jerarquía de DOS NIVELES" (sección A / subsección B; identidad del cuadro = (A, B, ocurrencia); fuente C → canal D → concepto E; plantilla autoridad; llenado dinámico) + plans/54 (lectura R1 única) + plans/46 (canal efectivo + sección por cierre) + plans/53 (fin del ordinal) + plans/50 §0.2 (fuentes opcionales) + plans/55 (resolvedor por etiquetas)
enmienda_jerarquia_ab: renombra el residuo R-SLOT-Q1-ENERBIT → R-JERARQUIA-AB (el nombre viejo embebía fuente y alcance; ver §7). Corrige el encuadre del documento anterior de este número (slots Q1 / alcance por ASE): queda invalidado y reemplazado completo, no parchado.
caso_rojo: corrida CLI real --periodo 2026101 --cinco-ase — RESULTADO ERROR codigo=ERR-SIN-SLOT salida=2 (fila fuente A|B|C|D|E='||ENERBIT|Total|' sin slot compatible en la sección esperada por el resolvedor vigente)
gate_entrada:
  plantilla_v2: NO SE TOCA (SHA 0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB — verificar intacto antes/despues)
  plantilla_v1: NO SE TOCA (historica; SHA 3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214 intacto antes/despues)
  alcance_aprobado: H0 read-only por bytes primero (plantilla × 3 periodos, 5 bloques × 3 periodos = 15 cruces); cero codigo de produccion en H0; prohibido ampliar sin nuevo OK del Ingeniero
  suite_base: p53-h2-suite.trx (1289/1222/67/0), p54-H3-suite.trx (1374/1307/67/0) y p55-H3-suite.trx (1484/1417/67/0) — el mismo conjunto de 67 rojos en las tres
  sin_commits: NUNCA git commit / git push; NUNCA git checkout / git reset (incidente plans/55 §8)
---

# Plan 56 — Jerarquía sección-subsección de la plantilla R1 (identidad de cuadro A/B)

> SDD completo PROPOSE → DESIGN → SPEC → TASKS. Fuente única de verdad para la implementación una vez aprobado ESTE documento. No se implementa sin aprobación del Ingeniero (punto de bloqueo: sin su OK no se abre H0; estado inicial PROPUESTO).
> Lectura obligatoria previa: `plans/55` (§7 residuo renombrado, §8 incidente `git checkout`) + `plans/54` (lectura única, CA0 `fuente × <f>`) + `plans/46` (canal efectivo + sección por cierre) + `plans/53` (fin del ordinal) + `plans/50` §0.2 (fuentes opcionales, invariantes fail-fast) + `.opencode/project-context.md` (doctrina madre, §"Jerarquía de DOS NIVELES").
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Mandato, corrección de encuadre y modelo (vinculante)

### 0.1 Mandato textual del Ingeniero (fuente única de este plan)

> *"Esto tampoco es que se aplique para Q1, así como no es de un ASE tampoco es de una quincena en particular, es concepto general del manejo de la plantilla."*
> *"No debe quedar quemado o que es de ASE5; esto si notas es de todos, no pertenece a un ASE, pertenece a un concepto de la plantilla."*

Consecuencia normativa: **cero identificadores de bloque, cero nombres de empresa, cero quincenas y cero direcciones congeladas en el camino de resolución.** La jerarquía es una propiedad de **la plantilla**: aplica a los 5 bloques y a ambas quincenas. Los casos concretos que siguen son **muestras de verificación**, nunca alcance.

### 0.2 Corrección explícita del encuadre anterior (este documento reemplaza, no parcha)

El contenido previo de este número 56 describía el defecto como "slots fuente-por-cuadro de Q1" con arbitraje de criterios de alineación (a)/(b)/(c) y veredictos de exceso. **Ese encuadre queda invalidado en su totalidad:**

1. El plan anterior afirmaba que la plantilla "no declara la sección" para la fila fuente y proponía reacomodar/arbitrar alineaciones entre cuadros. **Falso:** la plantilla sí declara dónde diligenciar; el error era de lectura (faltaba la dimensión A).
2. El plan anterior acotaba el alcance a una quincena y a un bloque ("Q1/ASE5/slots"). **Falso por mandato §0.1:** es concepto general de la plantilla.
3. El plan anterior trataba la fila-frontera de la columna B como "la sección". **Falso por doctrina:** esa frontera es la **subsección**, no la sección.

Nada de ese documento sobrevive: ni candidatos, ni arbitraje, ni veredictos de exceso, ni su §0. Ninguna implementación futura puede citarlo.

### 0.3 Modelo A/B (doctrina madre, `.opencode/project-context.md` §"Jerarquía de DOS NIVELES")

| Nivel | Columna | Valores | Qué cierra |
|---|---|---|---|
| **Sección** | **A** | `AFaseo`, `Componente`, `Subs/Cont` (`Total` = gran total) | el bloque completo de la sección |
| **Subsección** | **B** | `Mes`, `Aplicacion nuevos x reversion` (y variantes de rótulo) | el sub-bloque dentro de la sección |

- Las etiquetas se escriben **al cierre** del bloque (fila de total): `B=<subsección>/Total` cierra el sub-bloque; `A=<sección>/Total` cierra la sección.
- **Identidad del cuadro = (sección A, subsección B, ocurrencia de B dentro de A).**
- Dentro de la subsección: **fuente (C) → canal (D) → concepto (E)**; los valores son **componentes** (encabezado de columna, resueltos por encabezado; nunca por posición).
- La plantilla es la autoridad; el llenado es dinámico: lo que la fuente no trae queda en el **cero del template**.

### 0.4 El defecto actual (medido, no supuesto)

Corrida real (`--periodo 2026101 --cinco-ase`) llega hasta la escritura y muere:

```text
RESULTADO ERROR codigo=ERR-SIN-SLOT salida=2
[ERR-SIN-SLOT] ASE 5: la fuente R1 trae la fila A|B|C|D|E='||ENERBIT|Total|'
en la sección esperada '1.ª sección' (ocurrencia #1) sin ningún slot compatible…
```

**Causa raíz:** el resolvedor vigente identifica "sección" con la **fila-frontera de la columna B** (`ContratoSeccionPeriodFreeR1.SeccionPorCierre`) y alinea por orden (`Alinear`). **Ignora la columna A.** Así, el primer bloque de la fuente se empareja con la sección `AFaseo` del template, donde esa fuente no está declarada — cuando su identidad real es `(Componente, Mes)` y su destino existe.

### 0.5 Evidencia como muestra (no alcance)

Muestra 1 — hoja `Reporte Componentes R1` v2, un bloque (rango medido por bytes): `AFaseo` 602–613 (`Mes` 613) · `Componente` 615–648 (`Aplicacion nuevos x reversion` 629, `Mes` 648) · `Subs/Cont` 650–674 (`Aplicacion` 660, `Mes` 674). Fuente del período: bloque 4–22 cierra en `Mes/Total` fila 23 y rotula `A=Componente` en 24 ⇒ **(Componente, Mes)**; bloque 25–37 cierra en 38 y rotula `A=Subs/Cont` en 39 ⇒ **(Subs/Cont, Mes)**. Con la identidad correcta, la fila fuente de la muestra cae en `C637` — dentro de `(Componente, Mes)`, filas 630–648 — **donde sí hay destino**. Cierre aritmético: bloque1→`F648`, bloque2→`F674` ⇒ `TOT_OPT = F674+F648+F613−L613−L648 = 7.544.253.279,94` = **R10 oficial**.

Muestra 2 — otro bloque, mismo modelo: sus 3 bloques caen en `(AFaseo, Mes)`, `(Componente, Mes)` y `(Subs/Cont, Mes)` ⇒ `TOT_OPT = 8.811.472.160,34` = **R10 oficial**.

Lo que estas muestras prueban es el **modelo**, no un período: el T0 (§2-D1) lo demuestra en **15/15** (5 bloques × 3 períodos: julio-1, julio-2, octubre-1) por la **misma ruta**, sin listas por bloque y sin fallback posicional.

---

## 1. PROPOSE

### 1.1 Intent

Que el resolvedor R1 lea la **jerarquía de dos niveles de la plantilla** —identidad de cuadro `(A, B, ocurrencia)` + cruce `(C, D, E)` por encabezado— de modo que **toda** fila fuente con destino declarado resuelva por la misma ruta en los 5 bloques y ambas quincenas, sin quemar bloque/empresa/quincena/dirección, sin fallback posicional, sin insert/delete, sin tocar `<f>`, y con el `DetRetri` = R10 oficial en los 3 períodos **sin mover un peso**.

### 1.2 In Scope (H0 + H1 + H2 + H3, prohibido ampliar sin nuevo OK del Ingeniero)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **H0/T0** | Mapa por bytes read-only (cero código de producción): secciones (A) y subsecciones (B) declaradas por la plantilla y presentes en la fuente, con ocurrencia dentro de la sección, rango de filas y etiqueta de cierre; cruce `(A, B, ocurrencia, C, D, E)` resuelto en 15/15 sin fallback posicional y sin listas por bloque; cierre aritmético contra el R10 de cada período al peso | Fixture verde + doc de evidencia; puerta H0→H1 |
| **H1** | Resolvedor jerárquico en `ContratoSeccionPeriodFreeR1` + consumo en `EscritorValoresSlotR1` (objeto autorizado; ver §2-D2) | CA-JER-1..CA-JER-4 verdes en corrida real |
| **H2** | Cobertura: 15 tests (bloque × período) + tests de la regla de destino inexistente + candado anti-quemado (dos capas) | Suite verde + set-diff exacto |
| **H3** | Doctrina A/B en `.opencode/project-context.md` (solo lo probado) + cierre de este plan | Diff de docs citado; cero producción |

### 1.3 Non-goals / out (vinculante)

- No tocar la plantilla (SHAs v1/v2 antes/después; cero `<f>`, cero `InsertRow`/`DeleteRow`).
- **Prohibido tocar** `PreparadorBasePipeline`, `AnclasExtempR1` y el recorte `ProcesadorPeriodo.cs:198-217` (`R-RECORTE-Q2` sigue sin levantar; `R-SIGMA-Q2` y `R-ANCLA-EXTEMP` intactos).
- Prohibido: insert/delete, `<f>`, remap, colocación por posición, reacomodo de una fila a otra sección, fallback posicional bajo ninguna forma (ni cursor por firma sin sección, ni "próximo slot libre", ni proximidad/orden físico).
- No corregir EXTEMP/R2 ni reclasificar divergencias ajenas (Planes 49/50/53 y sus deudas intactas).
- No reabrir la semántica de agregados (Plan 25 D-C/D-F), el contrato por operando (Plan 53), la lectura única R1 (Plan 54) ni el resolvedor por etiquetas (Plan 55).
- Commits, UI, CLI, docs de usuario. NUNCA `git commit` / `git push`. **NUNCA `git checkout` / `git reset`** (incidente `plans/55` §8: trabajar siempre sobre copia previa del archivo).

### 1.4 Riesgos con mitigación

| ID | Riesgo | Mitigación |
|---|---|---|
| R-QUEMADO | Fijar el caso como bloque/fuente/quincena concretos (lista por bloque, literal de empresa, rama por quincena, dirección congelada) | Candado de dos capas §2-D5 + CA-JER-1 + G-D-H extendido (grep de cierre) |
| R-FALLBACK | Reintroducir colocación por posición por la puerta de atrás ("slot más cercano", "siguiente libre", alinear por orden sin A) | H0 exige resolución 15/15 sin fallback; candado conductual (reorden-dentro-de-subsección coloca idéntico); CA de formato |
| R-REACOMODO | "Arreglar" el rojo moviendo la fila a otro cuadro/sección (remap) en vez de leer su identidad | Prohibición expresa §1.3 + H1 solo consume identidad leída; fila sin destino ⇒ §2-D4, nunca reacomodo |
| R-DETRI | El cambio mueve el `DetRetri` (el dinero final) | CA-JER-3: `DetRetri` = R10 al peso en los 3 períodos; mismatch ⇒ STOP |
| R-ORACULO | Usar `<v>` stale de celdas con `<f>` como oráculo (lección Planes 28/29) | Método H0: texto-`<f>` + literales + fuente leída; `<v>` de celda con `<f>` nunca es oráculo (S0 G-WORKBOOK) |
| R-PLANTILLA | El fix toca una `<f>` o la geometría | SHAs + G-FORMULA + guard `ERR-PLANTILLA` intacto; H0 read-only |
| R-SILENCIO | La fila sin destino se descarta en silencio o se inventa celda | §2-D4: error accionable + alto, sin escribir nada; cero descarte silencioso, cero celda inventada |

---

## 2. DESIGN

### D1. Método H0/T0 por bytes (read-only bloqueante, cero código de producción)

**Forma normativa:** BCL puro (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`), sin Excel/COM, sin OpenXML SDK, sin ExcelDataReader. Rutas explícitas a los tres períodos (julio-1, julio-2, octubre-1); prohibido `Insumos.Raiz()` en lo nuevo. Cero escritura sobre fuentes y plantillas. Cero código de producción en H0: solo fixture T0 + documento de evidencia `Docs/T0-Plan56-Jerarquia-AB.md`.

Por cada cruce (bloque × período, 15 en total):

1. **Secciones y subsecciones del template**: por bloque, partición por columna A (sección) y por filas-frontera de columna B (subsección), con ocurrencia de B dentro de A, rango de filas y etiqueta de cierre (`B=<sub>/Total`, `A=<sec>/Total`).
2. **Secciones y subsecciones de la fuente**: misma partición leída de la fuente (el rótulo `A=<sección>` viaja en la fila de cierre; la subsección en la fila-frontera `B=<sub>/Total`).
3. **Cruce**: por cada fila fuente, identidad `(A, B, ocurrencia, C, D, E)` → celda declarada por el cuadro; componentes por **encabezado**, nunca por posición.
4. **Cierre aritmético**: `TOT_OPT` evaluado (texto-`<f>` + literales, nunca `<v>` stale) vs agregado del dominio vs R10 del período, al peso.

**Puerta H0→H1:** 15/15 resuelto por la misma ruta, sin fallback posicional, sin listas por bloque; cierres aritméticos al peso. Si algún cruce no resuelve ⇒ STOP + escalado (no se inventa celda, no se reacomoda, no se implementa H1).

| Candidata | Veredicto | Motivo |
|---|---|---|
| **M-A: T0 bytes A/B cuadro-por-cuadro en 15/15 + identidad (A, B, ocurrencia)** | **Elegida** | Es el único método que distingue "lectura sin dimensión A" de "falta real de destino" sin tocar nada; precedente: T0-Plan50, T0-Plan52, Plan 53 D2, Plan 55 T0 |
| M-B: parchar el `Alinear` para el caso muestra | Descartada | Es lista por bloque + literal embebido; viola el mandato §0.1 y R-QUEMADO |
| M-C: relajar `ERR-SIN-SLOT` a warning y continuar | Descartada | Convierte un defecto de lectura en falso verde; viola E1-1/E1-2 (Plan 40) y R-SILENCIO |

### D2. Resolvedor jerárquico (decisión)

**Objeto autorizado (y único):** `Remuneracion.Core/Models/ContratoSeccionPeriodFreeR1.cs` (predicado de identidad + consumo de `Alinear` corregido) + su consumo en `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs` (agrupación por identidad y `TryPlanificar`). Todo lo demás intacto: `SeccionEspejoR1.PorFronteraArriba` y `R1FirmaInterior` (consumidor Plan 35-T2), el lector leaf único (Plan 54), el contrato por operando (Plan 53), el resolvedor por etiquetas (Plan 55).

**Forma normativa del cambio:**

1. La identidad del cuadro se lee como **(sección A, subsección B, ocurrencia de B dentro de A)** — la columna A manda en el nivel sección; la fila-frontera de B tipa la subsección (no la sección).
2. La llave de colocación pasa a ser **(sección A + subsección B + ocurrencia + firma C/D/E + canal efectivo)** con componentes por encabezado. El `Alinear` por orden sin A queda sustituido donde el H0 demuestre que falla; donde el H0 demuestre identidad bajo ambas lecturas, el comportamiento se conserva byte-idéntico.
3. Geometría pura (D-H): el predicado conoce **etiquetas y ocurrencias**, nunca identificadores de bloque, nombres de empresa, quincenas ni direcciones congeladas (ver candado §2-D5).
4. La neutralidad por plan-then-write (Plan 45/46: bloque no colocable ⇒ comportamiento histórico) se conserva salvo que el H0 demuestre con bytes que la identidad nueva la subsume; cambiarla sin esa demostración ⇒ STOP.

### D3. Reglas de ausencia y de destino inexistente (normativas)

- **Ausencia en la fuente → cero del template, sin error.** Subsección, canal o fuente que el cuadro declara y la fuente no trae: la celda destino queda en el **0 del template**, con huella donde el mecanismo ya la emita. Nunca abortar por ausencia; los invariantes de estructura del reporte siguen fail-fast (Plan 50 §0.2).
- **Combinación presente en la fuente y ausente en todo el bloque de la plantilla → error accionable + alto, sin escribir nada.** No se reacomoda a otra sección, no se inventa celda, no se coloca por posición, no se descarta en silencio. El mensaje nombra **sección A + subsección B + fuente + canal + concepto + hoja + rango**, en vocabulario de la propietaria (componente/fuente/canal/concepto/cuadro), **sin culpar a la fuente**. El `ERR-SIN-SLOT` permanece como esta última puerta (formato de mensaje actualizado a la identidad A/B; la puerta sigue siendo fail-fast antes de cualquier escritura del bloque).

### D4. Lo que NO se toca (lista cerrada, verificable por grep de cierre)

`PreparadorBasePipeline`, `AnclasExtempR1`, el recorte `ProcesadorPeriodo.cs:198-217` (`R-RECORTE-Q2` sigue sin levantar; `R-SIGMA-Q2` vigente), `WorkbookLeafCellMapQ2`/`PorAse` (geometría), `ResolvedorConciliacionEmpresa`, `ValidadorBasico`/`WorkbookLeafCoherence` (juicio intacto; solo ganan cobertura si el H0 la ordena citar), tolerancia ±0,5 intacta, `<f>` de la plantilla intactas. **Cambio autorizado (no es "tocar", es el objeto):** `ContratoSeccionPeriodFreeR1` (+ consumo en `EscritorValoresSlotR1`) bajo §2-D2/D3.

### D5. Candado anti-quemado (dos capas, nuevo)

Patrón Plan 54/55, adaptado a este plan:

1. **Estático:** el camino de resolución no contiene identificadores de bloque, nombres de empresa, ramas por quincena ni direcciones congeladas. Allow-list cerrada (predicado A/B + `SinonimosEmpresaR1` + columnas por encabezado + canal efectivo) + **prueba de injerto que debe enrojecer** (un id de bloque, un nombre de empresa o una dirección congelada injertados a propósito en el camino enrojecen el test).
2. **Conductual:** el mismo fixture resuelto para los 5 bloques produce las **mismas reglas** (misma identidad → misma celda), sin ramas: la pertenencia a un bloque no decide. Más reorden-dentro-de-subsección coloca idéntico (la posición física no decide).

Sin candado verde no hay H3.

### Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Identidad A/B = juicio de lectura por etiquetas (Core puro); Writer = pegado values-only (intacto salvo consumo); Validador = juicio (intacto); destino inexistente = error accionable, no rama |
| **OCP** | ✅ Sección/subsección futura absorbida por la identidad que el cuadro declara, nunca por rama ni por entrada quemada; celda nueva = etiqueta del template, nunca `if (bloque/empresa/quincena)` |
| **LSP/DIP** | ✅ Contratos `IRecaudoReader`/`IPlantillaWriter`/`IValidador` intactos; cambios detrás de las mismas abstracciones |
| **Best practices** | ✅ Contract-first (H0 congela identidad 15/15 antes del juicio); fail-fast con A+B+fuente+canal+concepto+hoja+rango; tolerancia ±0,5 solo en comparación; values-only; D-H (cero ids de bloque/empresa/quincena/direcciones en resolución); CRLF; sin commits; prohibido `git checkout`/`git reset`; prohibido `<v>` de celda con `<f>` como oráculo |
| **Performance** | ✅ H0 acotado BCL ZIP+XML + fixtures existentes; identidad O(filas × etiquetas) por bloque en memoria; regens a temp fresca; cero workbooks extra en runtime |

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda H)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa trx + **set-diff vs `p53-h2-suite.trx` (1289/1222/67/0), `p54-H3-suite.trx` (1374/1307/67/0) y `p55-H3-suite.trx` (1484/1417/67/0)**: altas/bajas una por una con causa citada; **rojo nuevo fuera de la lista H0 ⇒ STOP**. 0 `Skip`, 0 `Assert.True(true)`; re-anclajes solo con evidencia (lectura §S0 de Plan 55: 0 bajas netas de cobertura, set-diff de resultados + conteo total).
- G-SHA: v1 `3086AC31…C5214` + v2 `0F192AD2…7F2DAB` intactos antes/después de cada H.
- G-FORMULA: `<f>` byte-idéntica en todo el diff (cero `new CellFormula` fuera de `EscribirFormulaPreservando`, cero `InsertRow`/`DeleteRow` fuera de Phase 0).
- G-D-H: cero ids de bloque / nombres de empresa / ramas por quincena / direcciones congeladas en el camino de resolución (grep de cierre + candado §2-D5).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale) o la fuente ya leída; **cero workbooks extra en runtime**.
- G-TRAZA: cada error nuevo nombra A+B+fuente+canal+concepto+hoja+rango donde aplique.
- G-PROC: cero commits; **cero `git checkout` / `git reset`**; CRLF sin mezclar en lo tocado.
- G-H0-READONLY: H0 no contiene código de producción (solo fixture H0 + doc de evidencia); el diff de `Remuneracion.Core/` + `Remuneracion.Infrastructure/` en H0 es vacío salvo tests `*T0Plan56*` / `*JerarquiaAB*`.

### CA-JER — Criterios de aceptación (vinculantes; con derivación `fuente × <f>`, CA0 del Plan 54)

- CA-JER-1: **cero ids de bloque / empresas / direcciones congeladas en resolución.** Camino verificado por grep de cierre + candado estático verde (incluida su prueba de injerto en rojo).
- CA-JER-2: **15/15 por la misma ruta.** Los 15 cruces (bloque × período) resuelven por la identidad A/B sin fallback posicional y sin listas por bloque; el mismo fixture para los 5 bloques produce las mismas reglas (candado conductual verde).
- CA-JER-3: **`DetRetri` = R10 en los 3 períodos sin mover un peso.** Cada `DetRetri` con su derivación `fuente × <f>` v2 + hoja/celda (CA0); mismatch ⇒ STOP.
- CA-JER-4: **error accionable con A+B+fuente+canal+concepto+hoja+rango y cero escritura.** Combinación presente en la fuente y ausente en todo el bloque ⇒ alto antes de escribir (nada del bloque se escribe), mensaje en vocabulario de la propietaria, sin culpar a la fuente; test que lo pincha verde.
- CA-JER-5: **plantilla byte-idéntica salvo los 11 tokens de período.** SHAs v1/v2 intactos; `<f>` byte-idéntica salvo lo aprobado por Phase 0.

### S1 — Criterios H0/T0 (read-only, cero código de producción)

- S1.1: secciones (A) y subsecciones (B) por bytes — declaradas por la plantilla y presentes en la fuente — con ocurrencia dentro de la sección, rango de filas y etiqueta de cierre, en 15/15; congelado como fixture.
- S1.2: cruce `(A, B, ocurrencia, C, D, E)` resuelto en 15/15 sin fallback posicional y sin listas por bloque; componentes por encabezado.
- S1.3: cierre aritmético por cruce contra el R10 de cada período, al peso, con derivación `fuente × <f>`.
- S1.4: lista cerrada de combinaciones presentes en la fuente y ausentes en todo el bloque (destino inexistente), si las hay, cada una con su veredicto de mensaje §2-D3; cero combinaciones sin veredicto.
- S1.5: lista cerrada de Δ de suite esperados (por defecto: ninguno — el H0 es read-only; cualquier rojo nuevo en H0 ⇒ STOP). **Sin H0 verde no se abre H1.**

### S2 — Criterios H1 (implementación)

- S2.1: identidad A/B implementada en `ContratoSeccionPeriodFreeR1` (+ consumo `EscritorValoresSlotR1`) según §2-D2; D-H intacto (grep + CA-JER-1).
- S2.2: reglas §2-D3 implementadas (ausencia → 0; destino inexistente → error accionable + alto sin escribir nada, con A+B+fuente+canal+concepto+hoja+rango).
- S2.3: CA-JER-2 + CA-JER-3 + CA-JER-4 reproducidos en corridas reales de los 3 períodos (SHAs + G-FORMULA + 0 warnings).
- S2.4: archivos prohibidos intactos (§2-D4, grep de cierre); tolerancia y mensajes del gate intactos en formato.

### S3 — Criterios H2 (regresión + cobertura)

- S3.1: 15 tests bloque × período (identidad pinchada con insumos reales por ruta explícita) + tests de la regla de destino inexistente (mensaje + cero escritura) + candado §2-D5 verde en sus dos capas.
- S3.2: set-diff vacío de rojos vs las tres líneas base; los 3 períodos verdes por la misma ruta (CA-JER-2/CA-JER-3); oráculo v2 julio-2 intacto.
- S3.3: code-review independiente **APROBADO** (0 bloqueantes): spec compliance + calidad + greps D-H + SHAs.

### S4 — Criterios H3 (docs/cierre)

- S4.1: doctrina A/B en `.opencode/project-context.md` (solo lo probado: identidad, reglas, cobertura + candado).
- S4.2: este plan como fuente de verdad; desviaciones con fecha/motivo/evidencia; cero código de producción en H3.

---

## 4. TASKS (fases H0..H3 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero código de producción)

Fixture: `Remuneracion.IntegrationTests/JerarquiaABT0Plan56Tests.cs` (BCL ZIP+XML, rutas explícitas a julio-1/julio-2/octubre-1; prohibido `Insumos.Raiz()` en lo nuevo). Evidencia: `Docs/T0-Plan56-Jerarquia-AB.md`.

| ID | Tarea | Verifica |
|---|---|---|
| T0.1 | Secciones (A) y subsecciones (B) del template por bytes: partición A/B por bloque, ocurrencia de B dentro de A, rango de filas, etiqueta de cierre | S1.1 |
| T0.2 | Secciones y subsecciones de la fuente por bytes: misma partición leída de la fuente en 15/15 (rótulo A en la fila de cierre, subsección B en la fila-frontera) | S1.1 |
| T0.3 | Cruce `(A, B, ocurrencia, C, D, E)` → celda en 15/15, sin fallback posicional, sin listas por bloque, componentes por encabezado | S1.2 |
| T0.4 | Cierre aritmético por cruce (texto-`<f>` + literales vs dominio vs R10, al peso, con derivación `fuente × <f>`) | S1.3 |
| T0.5 | Combinaciones sin destino en todo el bloque: lista cerrada + veredicto de mensaje §2-D3 por cada una | S1.4 |
| T0.6 | Grep cobertura: el gate de identidad sin tests propios + lista cerrada de Δ de suite | S1.5 |
| T0.7 | SHAs v1/v2 antes/después de H0 + `<f>` de cierre R1 intactas | S0 G-SHA/G-FORMULA |

- **Puerta H0→H1:** fixture verde + S1.1-S1.5 cerrados + APROBADO del Ingeniero. Si algún cruce no resuelve o algún cierre no cuadra ⇒ STOP + escalado (no se abre H1).

### H1/F1 — Resolvedor jerárquico + reglas (traza a §2–§3)

| ID | Tarea | Traza |
|---|---|---|
| H1.1 | Identidad A/B en `ContratoSeccionPeriodFreeR1` (+ consumo `EscritorValoresSlotR1`): etiquetas A/B + ocurrencia + firma C/D/E + canal efectivo, D-H intacto, sin fallback posicional, sin remap, sin reacomodo | D2, S2.1 |
| H1.2 | Regla de ausencia → 0 + regla de destino inexistente (error accionable + alto sin escribir, con A+B+fuente+canal+concepto+hoja+rango) | D3, S2.2 |
| H1.3 | e2e 3 períodos (5 bloques): CA-JER-2 + CA-JER-3 + CA-JER-4 (cero `ERR-SIN-SLOT` donde hay destino; `DetRetri` = R10 al peso) | S2.3 |
| H1.4 | Greps de cierre §2-D4 + SHAs + G-FORMULA + 0 warnings | S2.4 |

- **Puerta H1→H2:** S2 verde + CA-JER-1..CA-JER-5 + 0w + SHAs + G-FORMULA + greps D-H.

### H2/F2 — Regresión, cobertura y evidencia

| ID | Tarea | Verifica |
|---|---|---|
| H2.1 | 15 tests bloque × período + tests de destino inexistente (mensaje + cero escritura) + candado §2-D5 (estático allow-list + injerto; conductual mismas-reglas + reorden-dentro-de-subsección) | CA-JER-1/CA-JER-2/CA-JER-4, S3.1 |
| H2.2 | Suite trx + set-diff vs las tres líneas base; fuera de lista = STOP | S0 G-SUITE |
| H2.3 | No-regresión 3 períodos misma ruta (CA-JER-2/CA-JER-3) + oráculo v2 julio-2 intacto | S3.2 |
| H2.4 | Code-review independiente (`code-reviewer`, este plan como spec): **APROBADO** requerido | S3.3 |

- **Puerta H2→H3:** set-diff exacto + CA-JER-1..CA-JER-5 + review aprobado.

### H3/F3 — Docs/cierre

| ID | Tarea | Verifica |
|---|---|---|
| H3.1 | Doctrina A/B en `.opencode/project-context.md` (solo lo probado) | S4.1 |
| H3.2 | Cierre de este plan (§8) con desviaciones fechadas si las hubo; cero código de producción | S4.2 |

**Orden sugerido:** H0 → H1 → H2 → H3. Regresión continua (suite + SHAs en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero); **prohibido `git checkout`/`git reset` sin copia previa** (incidente Plan 55 §8).

---

## 5. Riesgos (consolidado + gates por fase)

Ver §1.4 (R-QUEMADO, R-FALLBACK, R-REACOMODO, R-DETRI, R-ORACULO, R-PLANTILLA, R-SILENCIO).

- **Gate H0→H1:** H0 verde + S1.3/S1.4 sin STOP + APROBADO del Ingeniero.
- **Gate H1→H2:** S2 + CA-JER-1..CA-JER-4 + 0w + SHAs + G-FORMULA + greps.
- **Gate H2→H3:** set-diff exacto + CA-JER-2/CA-JER-3/CA-JER-5 + review APROBADO.
- **Gate H3:** diff de docs citado; cero producción.

---

## 6. Clarification Gate (requiere aprobación del Ingeniero; sin esto no se abre H0)

1. **H0 primero:** ¿se ratifica H0 read-only bloqueante (cero código de producción) en 15/15 (5 bloques × julio-1/julio-2/octubre-1), con la regla "cruce sin resolver o cierre que no cuadra ⇒ STOP"?
2. **Modelo D2:** ¿se ratifica la identidad del cuadro `(A, B, ocurrencia)` con llave de colocación (A + B + ocurrencia + C/D/E + canal efectivo, componentes por encabezado) en `ContratoSeccionPeriodFreeR1` + `EscritorValoresSlotR1`, con las prohibiciones de remap/reacomodo/posición/insert/delete?
3. **Reglas D3:** ¿ausencia en la fuente → cero del template sin error, y destino inexistente en todo el bloque → error accionable + alto sin escribir nada (A+B+fuente+canal+concepto+hoja+rango, sin culpar a la fuente)?
4. **Candado D5:** ¿candado anti-quemado de dos capas (estático allow-list + injerto; conductual mismas-reglas) con CA-JER-1 como criterio?
5. **Línea base:** ¿se ratifican `p53-h2-suite.trx` (1289/1222/67/0), `p54-H3-suite.trx` (1374/1307/67/0) y `p55-H3-suite.trx` (1484/1417/67/0) — mismo conjunto de 67 rojos — para el set-diff, con rojo fuera de lista = STOP?
6. **Puerta:** sin 1-5 no se abre H0; sin H0 verde no hay H1; sin candado ni review no hay H3.

Sin estas aprobaciones el plan no avanza de PROPUESTO a APROBADO, y sin APROBADO no se abre H0.

> **RATIFICADO 2026-10-23 por el Ingeniero (con calidad senior) — gate §6 completo (1-6):** H0 read-only bloqueante 15/15 · identidad `(A, B, ocurrencia)` + llave `(A + B + ocurrencia + C/D/E + canal efectivo)` con componentes por encabezado · ausencia→0 y destino inexistente→error accionable+alto · candado anti-quemado de dos capas · líneas base p53/p54-H3/p55-H3 · puertas H0→H1→H2→H3.
> **Condición del Arquitecto (vinculante):** el plan es un **concepto general de la plantilla** — cero ids de bloque/ASE, nombres de empresa o quincenas en el camino de resolución, y cero direcciones congeladas. Las muestras citadas son verificación, no alcance. **H0 cierra con la matriz 15/15 y se detiene**; la apertura de H1 (escritura) requiere ratificación del Ingeniero sobre esa matriz.

---

## 7. Residuos declarados (fuera de alcance — NO implementar)

**R-JERARQUIA-AB (vivo, objeto de este plan; antes `R-SLOT-Q1-ENERBIT`):** renombrado por mandato del Ingeniero — el nombre viejo embebía fuente y alcance y queda retirado. La corrida 5-bloques aborta en `ERR-SIN-SLOT` del camino espejo-values porque el resolvedor vigente ignora la columna A. **Gate de resolución:** este plan (identidad A/B + reglas §2-D3 + candado §2-D5). Sin remap, sin insert, sin colocación por posición.

**R-RECORTE-Q2 (vivo, sin tocar):** `Remuneracion.Core/Services/ProcesadorPeriodo.cs:198-217` mantiene la omisión de la conciliación por empresa en Q2. Deuda `R-MAPA-V2-R2-ASE4` (Plan 50 §8). Gate de reactivación: Plan 55 §D6. No se toca acá (§2-D4).

**R-SIGMA-Q2 (vivo, no se re-litiga):** el gate Sigma no es descomposición válida del visible en Q2 (T0-Plan50 §4.2). No se toca acá.

**R-ANCLA-EXTEMP (vivo, Plan 54 §7 / Plan 55 §7):** `AnclasExtempR1` congela `F682`/`F702` como literales mientras la v2 trae `F682 = <f>F660+F629-L629</f>`. Con su propio T0 y plan. No se toca acá (§2-D4).

**R-ETIQUETAS-SIN-DESTINO (marco aplicable, Plan 55 §7):** si el cruce de etiquetas encuentra un exceso real futuro, aplica la regla de destino inexistente §2-D3 (error accionable + alto, sin escribir); prohibido inventar celdas destino.

---

## 8. Estado

**APROBADO 2026-10-23** (gate §6 ratificado). H0 habilitado y **bloqueante**: cierra con la matriz 15/15 y se detiene. H1 (escritura) queda condicionada a la ratificación del Ingeniero sobre esa matriz. Padre: Plan 55 (CERRADO). Madre doctrina: `.opencode/project-context.md` §"Jerarquía de DOS NIVELES" + Plans 50 §0.2 / 53 / 54 / 55 + Plan 46 (familia de lectura por etiquetas).
