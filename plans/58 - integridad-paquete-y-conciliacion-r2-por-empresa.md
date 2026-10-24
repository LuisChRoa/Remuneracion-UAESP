---
titulo: Plan 58 - Integridad del paquete y conciliacion R2 por empresa (estructural, sin casos particulares)
estado: APROBADO 2026-10-23 — P0 ratificado por el Ingeniero ("lo que necesito es que se repare la App para que no genere este error"). P1 (EAB) y P2 (columna E) EN ESPERA de decisión de negocio: el `EAB` de -18.613,87 no existe en la v2, y la plantilla v2 no cuadra internamente (empresas 95.653.588,19 vs total de bloque 95.634.974,32).
fecha: 2026-10-23
numero: 58
planes_ocupados: 01..57 ocupados; este es el 58
padre: plans/57 - anclas-validacion-derivadas-de-plantilla.md
madre_doctrina: .opencode/project-context.md (plantilla autoridad; values-only; guard anti-formula; tolerancia ±0,5; prohibido <v> stale como oraculo) + plans/54 (CA0 derivacion obligatoria; lectura unica) + plans/55 (conciliacion por empresa, anti-tautologia, candado) + plans/56 (identidad A/B, §7/§8) + plans/57 (§8 y deuda R-ANCLA-EXTEMP) + plans/40 (W-1 evalua+sella) + plans/31 (gate workbook-vs-dominio)
caso_rojo: comparacion 2026101 contra el resultado oficial del administrativo — CONSOLIDADO/D104:D108/D109 y los 25 componentes coinciden al centavo, pero (P0) el workbook generado es corrupto (Excel repara eliminando la formula compartida de sheet16.xml/INTERVENTORIA), (P1) VALIDACION_TOTAL!D34 = FALSE por Δ 18.613,87 en REMUNERACION Reciprocidad EAAB ASE 2, y (P2) la columna E de disposicion final esta vacia en R2/R4
gate_entrada:
  plantilla_v2: NO SE TOCA (SHA 0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB — verificar intacto antes/despues)
  plantilla_v1: NO SE TOCA (historica; SHA 3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214 intacto antes/despues)
  alcance_aprobado: H0 read-only por bytes primero (tres frentes, cero codigo de produccion); prohibido ampliar sin nuevo OK del Ingeniero
  suite_base: Remuneracion.IntegrationTests/TestResults/p57-suite.trx (1544/1509/35/0) — rojo fuera de lista = STOP
  sin_commits: NUNCA git commit / git push; NUNCA git checkout / git reset (hay trabajo sin commitear de 54/55/56/57)
  excel_cerrado: Excel debe estar cerrado durante las corridas
---

# Plan 58 — Integridad del paquete y conciliación R2 por empresa

> SDD completo PROPOSE → DESIGN → SPEC → TASKS. Fuente única de verdad para la implementación una vez aprobado ESTE documento. No se implementa sin aprobación del Ingeniero (punto de bloqueo: sin su OK no se abre H0; estado inicial PROPUESTO).
> Lectura obligatoria previa: `plans/57` (§8 y deuda) + `plans/56` (§7/§8) + `plans/55` (conciliación por empresa) + `plans/54` (CA0) + `.opencode/project-context.md` (doctrina madre).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Mandato vinculante, evidencia medida y lo que ya coincide (no tocar)

### 0.1 Restricción vinculante del Ingeniero (textual, norma suprema de este plan)

> "no debemos resolver casos de forma particular los ajustes deben ser a nivel de procesamiento si es algo que debemos mejorar, esto para que funcione la octubre-1 y después se pruebe otra y otro ajuste particular, nada de eso a nivel de proceso"

Interpretación obligatoria:

1. **Prohibido** condicionar por hoja, celda, ASE, empresa, quincena o período (`if (hoja == "INTERVENTORIA")`, `if (ase == 2)`, `if (empresa == "EAB")` y toda forma equivalente, incluida la que se esconda en un mapa, una allow-list o un mensaje).
2. **Prohibido** sumar/restar un delta constante para "cerrar" (18.613,87 o cualquier otro; ni siquiera con signo contrario ni repartido).
3. Cada corrección debe ser una **regla derivada de la estructura de la plantilla y de la fuente**, válida para cualquier período.
4. Si el diseño no puede evitar un caso particular ⇒ **STOP y explicarlo**; no se implementa el parche.

### 0.2 Archivos de evidencia (NO modificar)

- Oficial (administrativo): `Docs\Prueba Octubre-1\Resultados\Resultado administrativo\Remuneracion 202610-1 Total.xlsx` — **nuevo oráculo de negocio de 2026101** para este plan.
- Nuestro: `Docs\Prueba Octubre-1\Resultados\3\Remuneración 202610-1 Total (sin validar).xlsx`.

### 0.3 Lo que YA coincide al centavo (prohibido tocar)

`CONSOLIDADO!D104:D108` = **14.842.425.135,77 / 11.302.637.604,96 / 8.823.787.300,37 / 10.483.172.762,09 / 7.550.668.664,99** y `D109` = **53.002.691.468,18**. También coinciden los 25 componentes: TOT_OPT `D9:D13`, R2 `D28:D32`, EXTEMP `D47:D51`, R4 `D66:D70`, AJUSTES `D85:D89`. Cualquier cambio que mueva un peso de estos valores ⇒ STOP.

### 0.4 P0 (bloqueante) — el workbook generado es corrupto

- Excel responde "Hemos encontrado un problema con contenido…" y repara **eliminando la fórmula compartida de `/xl/worksheets/sheet16.xml`** (hoja INTERVENTORIA).
- Plantilla v2: `R26 = <f t="shared" ref="R26:R31" si="0">J15</f>` (maestro del grupo).
- Nuestro archivo: `R26 = <f>N26</f>` — perdió `t="shared"`, `ref` y `si`; `R27:R31` conservan `t="shared" si="0"` ⇒ **grupo huérfano sin maestro**.
- Restaurar `sheet16.xml` desde el template hace que Excel abra sin reparación (verificado por el Arquitecto): el daño lo produce **nuestra escritura sobre una hoja que el propio log declara "insumo externo declarado — hoja intacta, no se escribe"**.
- Efecto numérico adicional: `J15 = 189.185.988` vs `N26 = 189.185.987` ⇒ la fórmula que pusimos vale **1 peso menos**.
- El `OpenXmlValidator` **no detecta** esta clase de defecto (solo reportó `pageSetup/@verticalDpi=0`, herencia del template que Excel tolera). Ningún test actual lo detecta.

### 0.5 P1 — `VALIDACION_TOTAL!D34` = FALSE (Δ 18.613,87)

- Nuestro `D34` falla; el del administrativo es `TRUE`.
- Origen localizado: `REMUNERACION Reciprocidad EAAB`, ASE 2, `D29` → `Rem. Anticipos R2'!E158` = **11.780.995,14** vs **11.762.381,27** del administrativo (que referencia `E149`). Δ = **18.613,87**.
- En el R2 del administrativo, dentro del bloque ASE 2, existe un sub-bloque etiquetado `EAB` con `Total = -18.613,87`; en nuestro archivo esa celda equivalente está en **0**.
- Las otras 4 hojas por empresa coinciden exactas: `REMUNERACION_ENEL` **50.322.380.909,65**; `REMUNERACION_ENERBIT` **139.510.580**; `REMUNERACION_OCCIDENTE_ Directa` **2.527.193.143,92**; `REMUNERACION_EAAB-CL` **0**.
- **Dato duro medido:** la fórmula de la fila RECIPROCIDAD difiere entre las dos plantillas. Plantilla v2 (la nuestra): `E158 = E115+E90+E70-K70-K90`. Archivo del administrativo: `E149 = E106+E94+E80+E65-K65-K80`. **NO asumir que "falta un término"**: el H0 debe medir, celda por celda y por etiqueta, qué valor escribimos en cada insumo del bloque y cuál es el que no cuadra. Prohibido compensar con el delta.
- El administrativo cumple la propiedad que nosotros no: **las empresas suman exactamente el total del bloque**. Esa es la propiedad de negocio a garantizar, no el número.

### 0.6 P2 — columna E "Recaudo por disposición final" vacía en R2 y R4

- En R1 coincide exacta en los 5 ASE. En R2 (`E28:E32`) y R4 (`E66:E70`) está en 0, arrastrando `E104:E109`. Δ por ASE: **−1.802.859,92 / −5.621.540,82 / −926.712,26 / +764.429,99 / −381.253,19**.
- Es informativa (no entra en el pagable) pero se está perdiendo el dato: hay que derivarla de la fuente y de la estructura de la plantilla, no de literales por hoja.

---

## 1. PROPOSE

### 1.1 Intent

Que el workbook generado **abra en Excel sin reparación y sin perder fórmulas**, que la conciliación R2 por empresa **cierre contra el total del bloque por construcción estructural** (las empresas suman el bloque en los 5 ASE, `D34 = TRUE`), y que la columna de disposición final de R2/R4 se pueble desde la fuente — todo por **reglas de procesamiento derivadas de la plantilla**, con cero casos particulares, candado anti-parche verde y regresión del administrativo como oráculo.

### 1.2 In Scope (H0 + H1 + H2 + H3, prohibido ampliar sin nuevo OK del Ingeniero)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **H0/T0** | Inventario read-only por bytes, cero código de producción, de los tres frentes (§2-D0): (a) quién escribe `<f>` + conjunto de celdas con `<f>` que hoy reciben escritura; (b) barrido workbook-wide de grupos de fórmula compartida huérfanos; (c) mapa etiqueta→celda + valor de cada insumo del bloque ASE 2 en plantilla v2 / nuestro output / oficial | Fixture verde + doc de evidencia; puerta H0→H1 |
| **H1** | P0 estructural (el escritor no puede tocar celdas con `<f>` + gate `ERR-INTEGRIDAD` post-escritura) + P1 estructural (conciliación por empresa reconciliada con el total del bloque) + P2 estructural (disposición final R2/R4 derivada) | CA-A1..CA-A5 verdes en corrida real |
| **H2** | Cobertura: candado anti-parche (estático + injerto) + gate de integridad reutilizable + regresión de negocio contra el administrativo (comparación semántica por etiqueta) | Suite verde + set-diff exacto |
| **H3** | Doctrina en `.opencode/project-context.md` (solo lo probado) + cierre de este plan | Diff de docs citado; cero producción |

### 1.3 Non-goals / out (vinculante)

- No tocar la plantilla (SHAs v1/v2 antes/después; cero `<f>`, cero `InsertRow`/`DeleteRow`, cero geometría).
- No tocar fórmulas ni geometría del workbook generado salvo lo que el diseño estructural autorice con cita (por defecto: nada — values-only).
- No corregir EXTEMP/R1 ni reclasificar divergencias ajenas (Planes 49/50/53 y sus deudas intactas; `R-RECORTE-Q2`, `R-SIGMA-Q2`, `R-ANCLA-EXTEMP` intactos salvo que el H0 demuestre con archivo:línea que viven en el camino tocado).
- No reabrir la semántica de agregados (Plan 25 D-C/D-F), la lectura única R1 (Plan 54), el resolvedor por etiquetas (Plan 55) ni la jerarquía A/B (Plan 56).
- No re-congelar tablas por celda/ASE/empresa/quincena/período bajo ninguna forma (ni mapa, ni allow-list, ni "solo lectura").
- Commits, UI, CLI, docs de usuario. NUNCA `git commit` / `git push`. **NUNCA `git checkout` / `git reset`** (hay trabajo sin commitear de 54/55/56/57).

### 1.4 Riesgos con mitigación

| ID | Riesgo | Mitigación |
|---|---|---|
| R-PARCHE | "Arreglar" P0/P1/P2 con un condicional o un delta por caso (el anti-patrón que el Ingeniero prohíbe) | §0.1 + §2-D5 (candado) + CA-A7; caso no derivable ⇒ STOP, no parche |
| R-TAUTOLOGIA | La conciliación se define como "suma de lo escrito" y queda trivialmente verde | §2-D3: agregado derivado de la fuente + propiedad empresas==bloque; test anti-tautología |
| R-COMPENSA | Sumar/restar 18.613,87 (o repartirlo) para que D34 dé TRUE | Prohibición expresa §0.1 + CA-A2 (cierra por residuo designado por la plantilla, nunca por constante) + test que pincha la compensación |
| R-ORACULO-V | Usar cachés `<v>` stale de celdas con `<f>` como oráculo | Método H0 y gates: texto-`<f>` + literales + fuente leída; `<v>` de celda con `<f>` nunca es oráculo |
| R-VALIDATOR-CIEGO | Confiar en `OpenXmlValidator` como prueba de integridad (no ve shared huérfanas) | Gate propio por bytes §2-D2 reutilizable; el validator queda como señal secundaria, nunca como gate |
| R-DIRECCION | Comparar contra el administrativo por dirección de celda (su bloque R1 del ASE 1 cierra en fila 60, el nuestro en 77) | Comparación semántica por etiqueta §2-D6; dirección literal en un comparador ⇒ STOP |
| R-DETRI | El cambio mueve `D104:D108`/`D109`/`DetRetri` | CA-A3: byte-idénticos al centavo; mismatch ⇒ STOP |
| R-PLANTILLA | El fix toca una `<f>` o la geometría | SHAs + G-FORMULA + guard `ERR-PLANTILLA` intacto; H0 read-only |

---

## 2. DESIGN

### D0. Método H0/T0 por bytes (read-only bloqueante, cero código de producción)

**Forma normativa:** BCL puro (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`), sin Excel/COM, sin OpenXML SDK, sin ExcelDataReader. Rutas explícitas a la v2 + al oficial del administrativo + a nuestro output + a los tres períodos (2026071, 2026072, 2026101); prohibido `Insumos.Raiz()` en lo nuevo. Cero escritura sobre fuentes, plantillas y archivos de evidencia. Cero código de producción en H0: solo fixture T0 + documento de evidencia `Docs/T0-Plan58-Integridad-Conciliacion.md`.

Tres frentes, por bytes y sin código de producción:

a. **Quién escribe fórmulas.** Localizar el camino que convirtió `R26` en `<f>N26</f>`: qué escritor/mutador tocó `sheet16.xml`, con qué predicado decidió escribir ahí, y por qué el guard anti-fórmula no lo detuvo (cita archivo:línea del guard saltado y de la llamada que lo saltó). Inventariar el **conjunto completo de celdas con `<f>` que hoy reciben escritura** en la corrida 2026101 (¿solo INTERVENTORIA o hay más?). Cota vinculante: **no suponer una sola**; inventariarla.
b. **Grupos de fórmula compartida.** Barrido de **todas las hojas** del workbook generado y de la v2: agrupar por `si`, exigir exactamente un maestro (`t="shared"` con `ref`) por cada `si` con seguidoras, y listar todo grupo huérfano (hoja + `si` + celdas). El diagnóstico ya existe conceptualmente (Plan 33: invariante workbook-wide en `Reporte Componentes R1`); aquí se generaliza a workbook-wide y **debe quedar como gate reutilizable** (no como script de un solo uso).
c. **R2 por empresa (bloque ASE 2).** Mapa **etiqueta→celda** de los insumos del bloque en las tres fuentes (plantilla v2, nuestro output, oficial) + **valor de cada uno**, identificando el que no cuadra. Medir cada término de `E158 = E115+E90+E70-K70-K90` (v2) contra su contraparte de `E149 = E106+E94+E80+E65-K65-K80` (oficial) **por etiqueta, nunca por dirección**. Incluir el sub-bloque `EAB` (`Total = -18.613,87` en el oficial, 0 en el nuestro) como fila del mapa, no como conclusión anticipada.

**Puerta H0→H1:** los tres frentes cerrados + fixture verde. Si algún frente demuestra que la única salida es un caso particular ⇒ STOP + escalado (no se abre H1).

| Candidata | Veredicto | Motivo |
|---|---|---|
| **M-A: T0 bytes de tres frentes + fixes estructurales + gates reutilizables** | **Elegida** | Única que cumple §0.1 (reglas de proceso, no parches); precedente: T0-Plan50/52/54/55/56/57 |
| M-B: restaurar `sheet16.xml` desde el template tras escribir (copiar la hoja sana encima) | Descartada | Escondería el escritor que corrompe en vez de cerrarlo; el próximo período lo reintroduce; viola §0.1 |
| M-C: re-escribir `R26` con el `t="shared"` correcto tras el daño | Descartada | Reintroduce escritura sobre `<f>` por la puerta de atrás; viola values-only y el guard |
| M-D: sumar/restar 18.613,87 (o repartir el residuo por empresa) para que D34 dé TRUE | Descartada | Compensación prohibida por §0.1; convierte un defecto de reparto en falso verde |

### D1. P0 estructural (decisión)

1. El escritor **no debe poder tocar celdas con `<f>`**: el guard ya existe; el H0 demuestra **por qué se saltó** (qué camino no lo cruza: ¿mutador espejo? ¿sello? ¿recomposición? ¿escritura leaf con `omitirR1` parcial?) y H1 lo cierra **estructuralmente** (todo camino de escritura cruza el mismo predicado "celda con `<f>` efectiva ⇒ `ERR-PLANTILLA` + alto + cero escritura", sin excepciones por hoja).
2. La hoja declarada "intacta, no se escribe" debe serlo **por construcción** (el predicado la protege), no por log. Si el diseño vigente necesita escribir algún literal legítimo en una hoja con fórmulas, ese literal se autoriza **por naturaleza de celda leída** (destino literal declarado por el cuadro), nunca por nombre de hoja.
3. Prohibido "reparar después": ni copiar `sheet16.xml` sano encima, ni re-emitir el `t="shared"` perdido. El gate §2-D2 aborta **antes de entregar** el archivo.

### D2. Gate de integridad del paquete post-escritura (decisión, reutilizable)

Gate por bytes (sin Excel en el gate), **antes de entregar el archivo**, que aborta con `ERR-INTEGRIDAD` accionable (hoja + celda/índice + clase de defecto + acción) si detecta cualquiera de:

1. Grupos de fórmula compartida con `si` sin maestro (el caso P0 generalizado workbook-wide).
2. Celdas duplicadas dentro de una hoja (misma referencia dos veces).
3. Filas desordenadas (orden de `row` no monótono donde el formato lo exige).
4. Índices de `sharedStrings` fuera de rango.
5. Referencias a hojas inexistentes (nombres/hyperlinks/nombres definidos que apuntan a hojas que no existen).

El gate vive como componente reutilizable (no como test de un solo período) y corre en **toda corrida que genere workbook**. `OpenXmlValidator` queda como señal secundaria: su ceguera ante shared huérfanas está documentada y ningún criterio depende de él.

### D3. P1 estructural (decisión)

1. La conciliación por empresa debe **reconciliar con el total del bloque** (propiedad de negocio): las empresas suman exactamente el total del bloque en los 5 ASE, derivando el reparto **de la estructura de la plantilla** (el cuadro declara sus filas/empresas; el cruce por etiquetas las resuelve — maquinaria Plans 55/56, sin reinventarla).
2. Si el reparto por etiquetas deja un residuo (la fuente trae combinaciones que el cuadro no declara, o el cuadro declara slots que la fuente no llena), el residuo se asigna **al concepto que la propia plantilla designe** (el término que la `<f>` del bloque declara como cierre), **nunca por nombre de empresa ni por constante**. La forma exacta (qué término absorbe, con qué aritmética verificada) la fija el H0 con bytes; si el H0 demuestra que la plantilla v2 es **internamente inconsistente** (ninguna regla estructural cierra), se documenta y **se decide con el Ingeniero — no se compensa**.
3. Prohibiciones: compensar con el delta, crear el sub-bloque `EAB` a mano, copiar la fórmula del administrativo (`E149`) sobre la v2, o re-etiquetar filas para que "cuadre". El H0 mide; H1 deriva.

### D4. P2 estructural (decisión)

Escribir la columna de disposición final en R2 (`E28:E32`) y R4 (`E66:E70`) **derivándola de la fuente y de la estructura de la plantilla** (el componente por encabezado, la fila por etiqueta — mismo cruce que D3), no de literales por hoja ni de direcciones por ASE. El H0 fija la derivación exacta (qué columna/fila fuente alimenta cada celda, con bytes); si alguna celda no tiene fuente declarada ⇒ cero del template + registro (doctrina fuentes opcionales, Plan 50 §0.2), nunca celda inventada.

### D5. Candado anti-parche (decisión, dos capas)

Patrón Plans 54/55/56, adaptado a este plan:

1. **Estático:** test que **enrojece ante literales de celda/ASE/empresa/quincena/delta en los caminos tocados** (escritura, conciliación por empresa, disposición final, gate de integridad). Allow-list cerrada (cruce por etiquetas + predicados estructurales + `SinonimosEmpresaR1` + columnas por encabezado + identidad A/B) + **prueba de injerto que debe enrojecer** (un `if (ase == 2)`, un `if (empresa == "EAB")`, un `if (hoja == "INTERVENTORIA")` o un `+ 18613.87m` injertados a propósito en el camino enrojecen el test).
2. **Conductual:** el mismo fixture resuelto para cualquier período produce las mismas reglas (la pertenencia a un período/ASE/empresa no decide); más prueba de que una fórmula injertada con término falso dispara el gate con el Δ exacto.

Sin candado verde no hay H3.

### D6. Regresión de negocio: el administrativo como oráculo (decisión)

El archivo del administrativo es el **nuevo oráculo de 2026101** para `D104:D108`, `D109` y `VALIDACION_TOTAL!D34 = TRUE`. La comparación es **semántica (por etiqueta, nunca por dirección)**: su archivo viene de otra versión de plantilla (su bloque R1 del ASE 1 cierra en fila 60, el nuestro en 77), así que el comparador resuelve identidad de cuadro/concepto por etiquetas y compara valores ±0,5. Comparar por dirección ⇒ STOP (falso rojo/falso verde por geometría ajena).

### Architecture Validation Certificate

| Principio | Veredicto |
|---|---|---|
| **SRP** | ✅ Escritor = values-only sobre literales declarados (no puede tocar `<f>`); Gate de integridad = juicio por bytes reutilizable; Conciliación = reparto estructural por etiquetas; Comparador = juicio semántico |
| **OCP** | ✅ Período/empresa/hoja nuevos absorbidos por la regla estructural, nunca por rama ni por entrada quemada; celda nueva = etiqueta del template |
| **LSP/DIP** | ✅ Contratos `IPlantillaWriter`/`IValidador` intactos; campos aditivos o intactos; cambios detrás de las mismas abstracciones |
| **Best practices** | ✅ Contract-first (H0 congela tres frentes antes del juicio); fail-fast accionable (`ERR-INTEGRIDAD`/`ERR-PLANTILLA` con hoja+celda+clase+acción); tolerancia ±0,5 solo en comparación; values-only; §0.1 (cero casos particulares); CRLF; sin commits; prohibido `git checkout`/`git reset`; prohibido `<v>` de celda con `<f>` como oráculo |
| **Performance** | ✅ H0 acotado BCL ZIP+XML + fixtures existentes; gate O(celdas + grupos shared) en memoria sobre el paquete ya generado; regens a temp fresca; cero workbooks extra en runtime |

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda H)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa trx + **set-diff vs `p57-suite.trx` (1544/1509/35/0)**: altas/bajas una por una con causa citada; **rojo nuevo fuera de la lista H0 ⇒ STOP**. 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados, 0 tolerancias ampliadas; re-anclajes solo con evidencia.
- G-SHA: v1 `3086AC31…C5214` + v2 `0F192AD2…7F2DAB` intactos antes/después de cada H.
- G-FORMULA: `<f>` byte-idéntica en todo el diff (cero `new CellFormula` fuera de `EscribirFormulaPreservando`, cero `InsertRow`/`DeleteRow` fuera de Phase 0). **Cero `<f>` escritas por nosotros** (CA-A1).
- G-PARCHE: cero literales de celda/ASE/empresa/quincena/delta en los caminos tocados (grep de cierre + candado §2-D5).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale) o la fuente ya leída; **cero workbooks extra en runtime**; **sin Excel en el gate**.
- G-TRAZA: cada error nuevo nombra concepto + bloque + hoja + celda donde aplique (+ esperado vs obtenido donde aplique).
- G-PROC: cero commits; **cero `git checkout` / `git reset`**; CRLF sin mezclar en lo tocado; Excel cerrado durante las corridas.
- G-H0-READONLY: H0 no contiene código de producción (solo fixture H0 + doc de evidencia); el diff de `Remuneracion.Core/` + `Remuneracion.Infrastructure/` en H0 es vacío salvo tests `*T0Plan58*` / `*IntegridadPaquete*` / `*ConciliacionR2*`.

### CA — Criterios de aceptación (vinculantes; con derivación al centavo, CA0 del Plan 54)

- CA0: **todo número de criterio viaja con su derivación** (`fuente × <f>` v2 + hoja/celda). Un número sin derivación no se acepta como criterio.
- **CA-A1**: el workbook generado **abre en Excel sin reparación** y sin perder fórmulas (verificación por bytes del grupo shared + apertura limpia). Cero `<f>` escritas por nosotros. Derivación: gate §2-D2 verde + `R26` conserva `t="shared" ref="R26:R31" si="0"` y `R27:R31` con maestro.
- **CA-A2**: `VALIDACION_TOTAL!D34 = TRUE` en 2026101, y **las empresas suman el total del bloque en los 5 ASE**. Derivación: reparto estructural §2-D3 + residuo al concepto designado por la plantilla (citado con su `<f>`), nunca por empresa ni por constante. El 18.613,87 no aparece como literal en producción (grep).
- **CA-A3**: `D104:D108` y `D109` siguen exactamente igual (**14.842.425.135,77 / 11.302.637.604,96 / 8.823.787.300,37 / 10.483.172.762,09 / 7.550.668.664,99** y **53.002.691.468,18**) y `DetRetri` sigue igual a R10. Mismatch de un peso ⇒ STOP.
- **CA-A4**: columna E de R2 (`E28:E32`) y R4 (`E66:E70`) poblada y coherente con el administrativo por ASE (cierra los Δ −1.802.859,92 / −5.621.540,82 / −926.712,26 / +764.429,99 / −381.253,19 del arrastre `E104:E109`). Derivación: fuente × estructura §2-D4, por etiqueta.
- **CA-A5**: corridas CLI `2026101`, `2026071`, `2026072` con exit 0; sin `ERR-PLANTILLA`, `ERR-SIN-SLOT`, `ERR-VALIDACION`, `ERR-INTEGRIDAD`.
- **CA-A6**: plantilla byte-idéntica (v2 `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB`, v1 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).
- **CA-A7**: cero casos particulares — el candado estático verde y el injerto en rojo.

### S1 — Criterios H0/T0 (read-only, cero código de producción)

- S1.1: camino que escribió `<f>N26</f>` localizado con archivo:línea + explicación de por qué el guard se saltó + **inventario cerrado de celdas con `<f>` que reciben escritura** (cota: no suponer una sola).
- S1.2: barrido shared workbook-wide (v2 + nuestro output): todo `si` con seguidoras tiene exactamente un maestro con `ref`; lista cerrada de huérfanos (el P0 entre ellos).
- S1.3: mapa etiqueta→celda + valor por insumo del bloque ASE 2 en las tres fuentes + identificación del término que no cuadra (medición por etiqueta de `E158` vs `E149`, sin asumir "término faltante").
- S1.4: derivación de la columna E de R2/R4 (qué fuente alimenta cada celda, por etiqueta) o veredicto de "sin fuente declarada ⇒ 0 + registro".
- S1.5: lista cerrada de Δ de suite esperados (por defecto: ninguno — el H0 es read-only; cualquier rojo nuevo en H0 ⇒ STOP). **Sin H0 verde no se abre H1.** Si algún frente solo cierra con caso particular ⇒ STOP documentado.

### S2 — Criterios H1 (implementación)

- S2.1: todo camino de escritura cruza el predicado anti-`<f>` (el guard saltado cerrado); la hoja "intacta" lo es por construcción; CA-A1 encaminado.
- S2.2: gate `ERR-INTEGRIDAD` post-escritura implementado y reutilizable (las 5 clases de §2-D2); aborta antes de entregar.
- S2.3: conciliación reconciliada con el total del bloque (§2-D3); CA-A2 encaminado; cero literales de empresa/delta en producción. Si la v2 es internamente inconsistente ⇒ documentado + decisión con el Ingeniero, no compensación.
- S2.4: disposición final R2/R4 derivada (§2-D4); CA-A4 encaminado.
- S2.5: CA-A3 + CA-A5 + CA-A6 reproducidos en corridas reales (SHAs + G-FORMULA + 0 warnings).

### S3 — Criterios H2 (regresión + cobertura)

- S3.1: candado §2-D5 verde en sus dos capas (incluida prueba de injerto en rojo: celda/ASE/empresa/quincena/delta injertados enrojecen).
- S3.2: gate de integridad con tests por clase de defecto (shared huérfana, duplicada, fila desordenada, sharedStrings OOR, hoja inexistente) + test de que `OpenXmlValidator` no se usa como gate.
- S3.3: regresión de negocio contra el administrativo (D104:D108, D109, D34 = TRUE) por comparación semántica por etiqueta (test que pincha que comparar por dirección falla por construcción: fila 60 vs 77).
- S3.4: set-diff vacío de rojos vs `p57-suite.trx` salvo altas citadas; code-review independiente **APROBADO** (0 bloqueantes).

### S4 — Criterios H3 (docs/cierre)

- S4.1: doctrina en `.opencode/project-context.md` (solo lo probado: integridad del paquete, conciliación reconciliada, disposición final derivada, candado, oráculo administrativo).
- S4.2: este plan como fuente de verdad; desviaciones con fecha/motivo/evidencia; cero código de producción en H3.

---

## 4. TASKS (fases H0..H3 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero código de producción)

Fixture: `Remuneracion.IntegrationTests/IntegridadPaqueteT0Plan58Tests.cs` (BCL ZIP+XML, rutas explícitas a v2 + oficial + nuestro output + 2026071/2026072/2026101; prohibido `Insumos.Raiz()` en lo nuevo). Evidencia: `Docs/T0-Plan58-Integridad-Conciliacion.md`.

| ID | Tarea | Verifica |
|---|---|---|
| T0.1 | Frente (a): localizar el camino que escribió `<f>N26</f>` (archivo:línea del escritor + del guard saltado) + inventario cerrado de celdas con `<f>` que reciben escritura en 2026101 (cota: no suponer una sola) | S1.1 |
| T0.2 | Frente (b): barrido shared workbook-wide en v2 y en nuestro output (agrupar por `si`, exigir un maestro con `ref`); lista cerrada de huérfanos; prototipo del gate reutilizable como fixture | S1.2 |
| T0.3 | Frente (c): mapa etiqueta→celda + valor por insumo del bloque ASE 2 en v2 / nuestro output / oficial; medición término a término de `E158` vs `E149` por etiqueta; fila `EAB` incluida como dato | S1.3 |
| T0.4 | Derivación de la columna E de R2/R4 por etiqueta (fuente × estructura) o veredicto "sin fuente ⇒ 0 + registro" por celda | S1.4 |
| T0.5 | Comparador semántico por etiqueta (prototipo fixture): identidad de cuadro/concepto sin direcciones; demostración fila-60-vs-77 | S3.3 (anticipado) |
| T0.6 | Lista cerrada de Δ de suite + SHAs v1/v2 antes/después de H0 + `<f>` de cierre intactas | S1.5 |

- **Puerta H0→H1:** fixture verde + S1.1-S1.5 cerrados + APROBADO del Ingeniero. Si algún frente solo cierra con caso particular o compensación ⇒ STOP + escalado (no se abre H1).

### H1/F1 — Fixes estructurales (traza a §2–§3)

| ID | Tarea | Traza |
|---|---|---|
| H1.1 | Cerrar el camino que escribe `<f>`: todo camino de escritura cruza el predicado anti-`<f>` (celda con `<f>` efectiva ⇒ `ERR-PLANTILLA` + alto + cero escritura); la hoja "intacta" protegida por construcción | D1, S2.1 |
| H1.2 | Gate `ERR-INTEGRIDAD` post-escritura reutilizable (5 clases de §2-D2), antes de entregar el archivo | D2, S2.2 |
| H1.3 | Conciliación reconciliada con el total del bloque (§2-D3); residuo al concepto designado por la plantilla; cero literales de empresa/delta | D3, S2.3 |
| H1.4 | Disposición final R2/R4 derivada de fuente × estructura (§2-D4) | D4, S2.4 |
| H1.5 | e2e 3 períodos: CA-A1 + CA-A2 + CA-A3 + CA-A4 + CA-A5 (exit 0, sin los 4 errores) | S2.5 |
| H1.6 | Greps de cierre §0.1/§1.3 + SHAs + G-FORMULA + 0 warnings | S2.5 |

- **Puerta H1→H2:** S2 verde + CA-A1..CA-A7 encaminados + 0w + SHAs + G-FORMULA + greps anti-parche.

### H2/F2 — Regresión, cobertura y evidencia

| ID | Tarea | Verifica |
|---|---|---|
| H2.1 | Candado anti-parche (estático allow-list + injerto en rojo; conductual misma-regla) | CA-A7, S3.1 |
| H2.2 | Tests del gate de integridad por clase de defecto + test de ceguera del `OpenXmlValidator` | S3.2 |
| H2.3 | Regresión contra el administrativo (D104:D108, D109, D34 = TRUE) por comparación semántica + test fila-60-vs-77 | S3.3 |
| H2.4 | Suite trx + set-diff vs `p57-suite.trx` (1544/1509/35/0); fuera de lista = STOP | S0 G-SUITE |
| H2.5 | Code-review independiente (`code-reviewer`, este plan como spec): **APROBADO** requerido | S3.4 |

- **Puerta H2→H3:** set-diff exacto + CA-A1..CA-A7 + review aprobado.

### H3/F3 — Docs/cierre

| ID | Tarea | Verifica |
|---|---|---|
| H3.1 | Doctrina en `.opencode/project-context.md` (solo lo probado) | S4.1 |
| H3.2 | Cierre de este plan (§8) con desviaciones fechadas si las hubo; cero código de producción | S4.2 |

**Orden sugerido:** H0 → H1 → H2 → H3. Regresión continua (suite + SHAs en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero); **prohibido `git checkout`/`git reset` sin copia previa** (incidente Plan 55 §8). Excel cerrado durante las corridas.

---

## 5. Riesgos (consolidado + gates por fase)

Ver §1.4 (R-PARCHE, R-TAUTOLOGIA, R-COMPENSA, R-ORACULO-V, R-VALIDATOR-CIEGO, R-DIRECCION, R-DETRI, R-PLANTILLA).

- **Gate H0→H1:** H0 verde + S1.5 sin STOP + APROBADO del Ingeniero.
- **Gate H1→H2:** S2 + CA-A1..CA-A5 + 0w + SHAs + G-FORMULA + greps.
- **Gate H2→H3:** set-diff exacto + CA-A6/CA-A7 + review APROBADO.
- **Gate H3:** diff de docs citado; cero producción.

---

## 6. Clarification Gate (requiere aprobación del Ingeniero; sin esto no se abre H0)

1. **H0 primero:** ¿se ratifica H0 read-only bloqueante de tres frentes (quién escribe `<f>` + inventario de celdas con `<f>` escritas; barrido shared workbook-wide reutilizable; mapa etiqueta→celda del bloque ASE 2 en las tres fuentes), con la regla "solo cierra con caso particular ⇒ STOP"?
2. **P0:** ¿el escritor no puede tocar celdas con `<f>` por construcción (cerrar el camino saltado, no restaurar la hoja después) + gate `ERR-INTEGRIDAD` post-escritura por bytes con las 5 clases, sin Excel en el gate?
3. **P1:** ¿la conciliación reconcilia con el total del bloque por regla estructural (residuo al concepto designado por la plantilla, nunca por empresa ni por constante), con prohibición de compensar el 18.613,87, y con STOP + decisión con el Ingeniero si la v2 es internamente inconsistente?
4. **P2:** ¿la disposición final de R2/R4 se deriva de fuente × estructura (sin fuente declarada ⇒ 0 + registro, nunca celda inventada)?
5. **Candado y línea base:** ¿candado anti-parche de dos capas (estático + injerto en rojo que pincha celda/ASE/empresa/quincena/delta) + set-diff contra `p57-suite.trx` (1544/1509/35/0), con rojo fuera de lista = STOP?
6. **Oráculo:** ¿el administrativo es el oráculo de 2026101 para D104:D108, D109 y D34 = TRUE, comparado por etiqueta (nunca por dirección: fila 60 vs 77)?
7. **Puerta:** sin 1-6 no se abre H0; sin H0 verde no hay H1; sin candado ni review no hay H3.

Sin estas aprobaciones el plan no avanza de PROPUESTO a APROBADO, y sin APROBADO no se abre H0.

---

## 7. Residuos declarados (fuera de alcance — NO implementar)

**R-RECORTE-Q2 (vivo, sin tocar):** `Remuneracion.Core/Services/ProcesadorPeriodo.cs:198-217` mantiene la omisión de la conciliación por empresa en Q2. Gate de reactivación: Plan 55 §D6.

**R-SIGMA-Q2 (vivo, no se re-litiga):** el gate Sigma no es descomposición válida del visible en Q2 (T0-Plan50 §4.2).

**R-ANCLA-EXTEMP (vivo, Plans 54 §7 / 55 §7 / 56 §7-O3 / 57):** ver `plans/57` §8 y su deuda; no se toca acá salvo que el H0 demuestre con archivo:línea que vive en el camino tocado.

**O1/O2 del Plan 56 (vivas, con gate):** candado estático parcial (`TryPlanificar`/`MensajeSinSlotJerarquia`) y registro durable del cero documentado solo en log + memoria. No se tocan acá.

**R-ETIQUETAS-SIN-DESTINO (marco aplicable, Plan 55 §7 / Plan 56 §7):** exceso futuro ⇒ error accionable + alto sin escribir; prohibido inventar celdas destino.

---

## 8. Estado

**PROPUESTO 2026-10-23.** Pendiente aprobación del Ingeniero (Clarification Gate §6). Sin APROBADO no se abre H0. Padre: Plan 57. Madre doctrina: `.opencode/project-context.md` + Plans 54/55/56/57 (+ 40/31 para el oráculo).
