# Plan 16 — HU-16: INTERVENTORIA + filas L-Especiales menores (última hoja sin HU)

> **Historia:** cerrar la última hoja sin HU — `INTERVENTORIA` (bloque K25:N31 por ASE según período; costo de interventoría por ASE) — con su T0 propio valor-vs-fórmula celda por celda contra ambos canónicos, más el mapeo de las celdas L-Especiales de filas operando menores sin mapear (0 en Q1, riesgo stale futuro — follow-up HU-07). Integración al procesador multi-ASE + resumen Serilog + paridad CLI↔UI si aplica.
> **Rector (IRRENUNCIABLE):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 3 (3.1–3.3 cerradas por HU-14/HU-15; esta HU es el frente "hoja restante"), §6 arquitectura (capas, preservación de fórmulas, Serilog, lectura directa de fuente), §10 CA + tolerancia ±0.5. **Entran SOLO INTERVENTORIA + L-Especiales menores + integración mínima.**
> **Origen funcional:** `Detalle de plantilla.docx` en lo que diga de INTERVENTORIA (rango L25:N31 por ASE según período) y de la columna "Especiales"; `Prompt Maestro Vo.docx` en lo que diga del costo de interventoría por ASE. `Proceso de Recaudo.docx` **NO aplica** salvo cita directa que lo exija.
> **Continuidad:** HU-01..HU-15 cerradas (ruta leaf multi-ASE atómica single-write, mapas explícitos por `Ase.Id` + período, validador estricto, UI modo 5-ASE, CLI con `CodigosSalida` 0-5 y paridad CLI↔UI, Golden Capa A, build 0 warnings, tests 259/259 + harness 24/24 regresión ciega). Este plan NO reabre su semántica: **cero cambios de cálculo/escritura existentes; Q1/Q2 intactos por construcción.**
> **Goldens canónicos (no se re-fijan):** Q1 `Docs/Insumos/Remuneracion 202607-1 Total.xlsx`, Q2 `Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx`.
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-09

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. Las dos incógnitas reales —el origen del costo de interventoría (¿insumo externo sin fuente en insumos?) y el estatuto de las L-Especiales menores (¿operando real o cero estático?)— **no se asumen**: las dicta el T0 contra ambos canónicos (§4 Fase 0), con desenlaces soportados sin rebase (ver D2/D3). El apply espera aprobación explícita.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (lectura de código/planes/disco + lectura Excel propia):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | Hoja `INTERVENTORIA` existe en el golden Q1 con bloque por ASE en filas 25–31: fila 24 header (`VALOR OFICIAL MES \| ASE \| Segunda quincena \| Primera quincena`, cols K–N), filas 25–29 un ASE por fila (K = valor oficial mes, L = Id ASE, M = segunda quincena, N = primera quincena), filas 30–31 totales | Lectura Excel `INTERVENTORIA!K24:N32` Q1 | Base estructural del T0; el rango exacto L25:N31 del Detalle vive dentro de este bloque (K25:K31 = columna de contexto, no objetivo) |
| V2 | Valores Q1 del bloque: ASE1 378371975 (189185988/189185987), ASE2 533160511 (266580256/266580255), ASE3 309577071 (154788536/154788535), ASE4 240782166 (120391083/120391083), ASE5 257980891 (128990446/128990445); total 1719872614 (859936309/859936305). Partición por mitades ±1 (redondeo), invariante anual — no derivable de R1/R2/R4 | Lectura Excel Q1 | Hipótesis de trabajo (a confirmar por T0): **insumo externo anual** (costo de interventoría por ASE, Prompt Vo), sin fuente en `Docs/Insumos/`; si T0 lo confirma → desenlace D2(b): se declara, no se inventa |
| V3 | El bloque Q2 (`Remuneracion 202607-2 Total.xlsx` raíz) replica los mismos valores Q1 en `K24:N32` (misma tabla anual) | Lectura Excel Q2-raíz | Indicio de tabla anual estática independiente del período; T0-0.2 lo verifica contra el **canónico Q2** (`Plantilla 8 agos…`), no contra este archivo |
| V4 | Set L-Especiales **mapeado** (Plan 07 T0 §0.1/§0.3): L25 (ASE1), L113+L90 (ASE2), L217+L238 (ASE3), L357+L391 (ASE4), L478+L498 (ASE5) en `Reporte Componentes R1`; más L10/L223/L369 = 0 explícitos (EXTEMP sin Especiales) | `plans/07 - HU-07 T0 Evidencia.md` §0.1–§0.3 | Punto de partida cerrado: el T0 de esta HU enumera la columna L **completa** y la difiere contra este set; solo la diferencia es trabajo nuevo |
| V5 | Doctrina aplicable cerrada: mapa explícito por `Ase.Id` + período, prohibidos offsets (HU-07 D1, HU-12 D1); golden honesto Capa A (nunca caché de salida vs golden; A5 HU-11/HU-12); T0 bloquea el mapa (ninguna celda entra al código sin evidencia); fail-fast nombra ASE+reporte+celda; sin insert/delete de filas | Planes 07/11/12 | Método, no contenido: este plan lo reutiliza literalmente |
| V6 | `INTERVENTORIA` ya vive en el mapa **protegido** (planes 11 §0.2/§2.3 y 12 §0.3/§2.3: "entra al mapa protegido, no se toca") | Planes 11/12 | Baseline: hoy la hoja se preserva intacta; esta HU decide con evidencia si sigue protegida (D2b) o gana superficie escribible/validada (D2a) |
| V7 | Integración disponible: `ProcesadorPeriodo` (path Q1/Q2 por ASE, fail-fast ASE+reporte), Serilog por paso/ASE/hoja + `RunId` inyectable (HU-14/HU-15), `CodigosSalida` 0-5, paridad CLI↔UI in-process, regresión 259/259 + 24/24 | Planes 14/15 + contexto del orquestador | La integración HU-16 es aditiva sobre estos contratos; no los reabre |
| V8 | Ningún finder del locator cubre interventoría; ninguna fuente `*Interventoria*` existe en `Docs/Insumos/` (naming R1/R2/R4/banco/balance/notas/retribución listado en Propuesta §5.1, sin fila de interventoría) | Propuesta §5.1 + planes 07–12 (finders) | Refuerza V2: **no hay fuente conocida**; el T0-0.3 debe dictaminar origen antes de cualquier lectura/escritura |

**NO verificado (y por eso T0 es bloqueante, §4 Fase 0):**

1. Carácter valor-vs-fórmula celda por celda del bloque K25:N31 (+ filas 30–31) en **ambos canónicos** (¿M/N formuladas como mitad de K? ¿todo valor estático? ¿alguna con referencia externa?).
2. Origen del costo: ¿alguna celda del bloque referencia otra hoja del workbook, o es insumo externo puro sin fuente en insumos?
3. Si el bloque se desplaza "según período" (Detalle): ¿mismas filas 25–31 en Q2 canónico o distinto rango por ASE?
4. Enumeración completa de la columna L en `Reporte Componentes R1` por ASE en ambos canónicos: ¿qué celdas L numéricas existen fuera del set V4, con qué valor y con qué fórmula?
5. Si alguna L-menor no-cero en Q2 canónico cierra contra alguna fuente (¿operando real?) o es ajuste manual.
6. Si la hoja exige extender validaciones 2.7 ya cerradas (solo si la hoja lo exige, y con recorte declarado).
7. Paridad CLI↔UI "si aplica": ¿el CLI necesita flags nuevos o basta la paridad de valores existente?

> **Regla de hierro del plan:** ninguna celda de INTERVENTORIA ni de L-Especiales-menores entra al código sin pasar por T0. Si L25:N31 no tiene fuente o no cierra ±0.5: **NEEDS_CONTEXT con recorte, nunca invención**. Lo verificado arriba (V1–V8) sí es contratable desde el día uno.

### 0.2 Mapeo al Rector (in vs out)

**Entra porque §9 Fase 3 / §6 / §10 + documentos base lo piden ahora:**

| Rector / base | Qué cubre HU-16 |
|---|---|
| Detalle de plantilla (INTERVENTORIA L25:N31 por ASE según período) | T0 propio del bloque + lectura/escritura/validación según veredicto T0 |
| Prompt Maestro Vo (costo de interventoría por ASE) | Semántica del bloque: costo externo por ASE; si no hay fuente → se declara como tal |
| Detalle (columna "Especiales") + follow-up HU-07 | Mapeo de L-Especiales menores con veredicto T0 (operando vs cero estático) |
| §6 trazabilidad + §10 CA-6 | Integración al procesador multi-ASE + resumen Serilog (`Hoja = "INTERVENTORIA"`) |
| HU-15 paridad CLI↔UI | Paridad si aplica (sin flags nuevos salvo que T0+integración lo exijan) |

**Sale porque §9 lo asigna a HU-17 u otras HUs (EXPLÍCITO):**

- 3.4 manual + entrega (HU-17): instructivo Capa B, paquete, casos, criterio de pase.
- Cambios de cálculo/escritura existentes (Q1/Q2 intactos por construcción; ningún mapa HU-07..HU-12 se reescribe).
- Validaciones 2.7 ya cerradas: solo se extienden si la hoja lo exige (Req 5, con recorte declarado).
- DI framework, restyle UI, instalador/setup.
- **Insert/delete de filas:** prohibidos. Si la plantilla no alcanza, **fail-fast honesto como en HU-07**.

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **T0 bloqueante con doble protocolo** (§4 Fase 0): T0-INTERVENTORIA (bloque K25:N31 + totales, ambos canónicos, valor-vs-fórmula + origen) y T0-L-ESPECIALES (columna L completa R1 por ASE, ambos canónicos, dif contra set V4). Sin ambos dumps no se escribe ni una celda HU-16. |
| G2 | **INTERVENTORIA con doble desenlace sin rebase** (D2): (a) si T0 demuestra operandos editables con fuente o aritmética que cierra ±0.5 → se leen/escriben/valid­an en la misma pasada; (b) si es insumo externo sin fuente en insumos → se declara como tal: hoja protegida intacta + assert de presencia/estructura + `Log.Warning` documentado o ceros documentados según T0-0.4, **nunca valor inventado**. |
| G3 | **L-Especiales menores con doble desenlace sin rebase** (D3): (a) celda L con operando real que cierra contra fuente → entra al mapa editable por `Ase.Id`; (b) cero estático del proceso manual (gemelo de la decisión F178/F521 HU-07T0 §0.2) → no se escribe, se protege + test que fija el 0 contra ambos goldens. |
| G4 | **Q1/Q2 intactos por construcción:** dispatch por `NumeroQuincena` donde aplique; rama HU-16 inactiva sin sus insumos; 259/259 + 24/24 como red ciega. |
| G5 | **Una sola llamada de escritura** por proceso (mismo overload lista HU-07..HU-12). Sin segundo pase. |
| G6 | **Golden honesto Capa A extendida a HU-16** (§2.7, doctrina HU-06..HU-12); Capa B manual residual como acción del usuario (HU-17 la instruye). |
| G7 | **Matcheo estricto `Single` por `Ase.Id`** en todo gate nuevo; fail-fast nombra ASE + reporte/hoja + celda. |
| G8 | **Paridad CLI↔UI mínima:** si la integración no agrega flags, la paridad es igualdad de valores entre vías (patrón HU-15 §2.7); si exige flags nuevos, entran `OpcionesCli` + matriz de parseo (Req 6). |

---

## 1. PROPOSE

### 1.1 Intent

Cerrar la última hoja sin HU con el mismo rigor que HU-07/T0: dictaminar con evidencia qué es el bloque INTERVENTORIA por ASE (editable con fuente vs insumo externo declarado) y qué son las L-Especiales menores (operando vs cero estático), integrar lo que T0 demuestre al procesador multi-ASE con resumen Serilog, y certificarlo contra ambos canónicos — sin mover una coma del comportamiento Q1/Q2 existente.

### 1.2 In Scope

- Discovery T0 doble (§4 Fase 0) + mapas congelados solo con evidencia (`WorkbookLeafCellMapInterventoria` y/o extensión L-menores, según veredictos D2/D3).
- Modelos mínimos aditivos en Core si T0 demuestra superficie editable (patrón `AjustesSfT = null`: nullable/default que preserva HU-12 intacto cuando no aplica).
- Readers header/rango-driven de lo que T0 demuestre + finders de locator solo si existe fuente (si no existe fuente, no hay finder: D2b).
- Validación pre/post con mapa ampliado (INTERVENTORIA visibles, L-menores, resto 2.7 como protegidas).
- Gates nuevos solo por quincena/período que T0 justifique (Q1 intacto); extensión 2.7 solo si la hoja lo exige (Req 5).
- Procesadores: integración HU-16 al flujo por ASE con fail-fast que nombra ASE + hoja + celda.
- Golden Capa A HU-16 contra ambos canónicos (§2.7) + regresión 259/259 + 24/24.
- Resumen INTERVENTORIA por ASE en log/Serilog (`Hoja = "INTERVENTORIA"`, delta mínimo UI) + paridad CLI↔UI si aplica (G8).

### 1.3 Out of Scope

Todo §0.2 (manual/entrega HU-17, cambios de cálculo/escritura existentes, 2.7 salvo exigencia de la hoja, DI, restyle UI, setup). Además: motor Excel/COM en CI; relectura de TXT diarios de EFC; `System.CommandLine`; `.gitattributes`; tocar el harness auxiliar; reabrir cadena 2.5/2.6 o conciliación HU-08 en Q2.

### 1.4 Resultado de negocio

El Ingeniero ejecuta el modo 5 ASE sobre `REMUNERACION 2026071` y `REMUNERACION 2026072` con los canónicos; obtiene la salida con INTERVENTORIA correcta post-Excel (o declarada-insumo-externo con fail-fast/ceros documentados y hoja intacta, según veredicto T0) y L-Especiales menores mapeadas o fijadas-en-0; Q1/Q2 existentes bit-a-bit iguales; el log audita interventoría por ASE; las pruebas demuestran cada celda HU-16 = evidencia ±0.5 o estatuto declarado con test.

### 1.5 Base documental (origen funcional — citas por sección)

| Documento | Sección / instrucción | Qué aporta a HU-16 |
|---|---|---|
| `Detalle de plantilla.docx` | Inst. hoja `INTERVENTORIA` (rango L25:N31 por ASE según período) + inst. columna "Especiales" | Layout objetivo y semántica; base de ambos T0 |
| `Prompt Maestro Vo.docx` | Costo de interventoría por ASE (pegado/valores por ASE) | Qué tramos son valores externos candidatos vs calculados |
| Rector Propuesta | §9 Fase 3 (frente restante), §6 (preservación, lectura directa, trazabilidad), §10 CA + ±0.5 | Rector normativo |
| `plans/07 - HU-07 T0 Evidencia.md` | Protocolo 0.1–0.5 + decisión F178/F521 (cero estático) + mapa editable por ASE | **Patrón a replicar**: T0-L-Especiales es su espejo; D3(b) es gemela de F178/F521 |
| `plans/11/12` | Doctrina golden honesto Capa A, T0 bloqueante, PRs encadenados, mapa por Id + período | Doctrina y método vigentes |
| `Proceso de Recaudo.docx` | — | **NO aplica** salvo cita directa que lo exija |

---

## 2. DESIGN

### 2.1 Tablas verificadas (contratables desde el día uno)

**Bloque INTERVENTORIA Q1 (lectura Excel propia, `INTERVENTORIA!K24:N32`):**

| Fila | K (VALOR OFICIAL MES) | L (ASE) | M (Segunda) | N (Primera) |
|---|---|---|---|---|
| 24 | header | header | header | header |
| 25 | 378371975 | 1 | 189185988 | 189185987 |
| 26 | 533160511 | 2 | 266580256 | 266580255 |
| 27 | 309577071 | 3 | 154788536 | 154788535 |
| 28 | 240782166 | 4 | 120391083 | 120391083 |
| 29 | 257980891 | 5 | 128990446 | 128990445 |
| 30 | 1719872614 | — | 859936309 | 859936305 |
| 31 | 1719872614 | — | — | — |

Invariantes observables (a confirmar valor-vs-fórmula por T0): M+N = K por ASE (±1 por mitades); ΣK = 1719872614; ASE4 reparte exacto (sin ±1). Ningún valor deriva de R1/R2/R4 (órdenes de magnitud distintos: cientos de millones vs miles de millones).

**Set L-Especiales mapeado V4 (Plan 07T0, punto de partida — no se re-descubre):**

| ASE | Slots L mapeados (R1) | L = 0 explícitos |
|---|---|---|
| 1 | L25 ← Esp(Mes[0]) | L10 ← 0 |
| 2 | L113 ← Esp(Mes[1]), L90 ← Esp(Mes[0]) | — (EXTEMP estático 0) |
| 3 | L217 ← Esp(Mes[0]), L238 ← Esp(Mes[1]) | L223 ← 0 |
| 4 | L357 ← Esp(Mes[0]), L391 ← Esp(Mes[1]) | L369 ← 0 |
| 5 | L478 ← Esp(Mes[0]), L498 ← Esp(Mes[1]) | — (EXTEMP estático 0) |

> Tablas de valor-vs-fórmula del bloque (ambos canónicos), enumeración L-completa con dif contra este set, y veredictos D2/D3 las congela T0; este plan no las inventa.

**Contratos y piezas que NO cambian:**

| Elemento | Estado | Tratamiento HU-16 |
|---|---|---|
| Mapas HU-07..HU-12 (leaf Q1/Q2, AjustesSfT, DetRetri, validaciones) | Certificados | Intactos; HU-16 agrega mapa hermano o extensión aditiva, nunca parche con `if` |
| `ConsolidadoAse.TotalAse` / `GranTotal` (getters) | Existen | Sin cambio salvo que T0 demuestre que interventoría alimenta el total (entonces Req 2 lo cubre aditivamente) |
| `INTERVENTORIA` como protegida (planes 11/12) | Baseline | Punto de partida; D2 decide con evidencia si sigue protegida o gana superficie |
| `CodigosSalida` 0-5, `CodigoDe`, RunId, niveles Serilog | HU-14/HU-15 | Reutilizados; sin códigos nuevos |
| Regresión 259/259 + harness 24/24 | Red | Debe seguir verde sin tocar goldens |

### 2.2 Decisiones de arquitectura

| ID | Opción elegida | Descartada | Por qué |
|---|---|---|---|
| D1 | `WorkbookLeafCellMapInterventoria` (+ extensión L-menores): mapa hermano explícito por (`Ase.Id`, `NumeroQuincena`) — visibles T0 + editables T0 + protegidas | Parchear `WorkbookLeafCellMapPorAse`/`Q2` con ramas o offsets | OCP + Q1/Q2 intactos: el bloque puede desplazarse "según período" (Detalle); mezclarlo rompería lo certificado. Doctrina HU-07 D1 / HU-12 D1 |
| D2 | INTERVENTORIA doble desenlace T0: (a) editable con fuente/arimética que cierra → leer+escribir+validar en la misma pasada; (b) insumo externo sin fuente → **protegida intacta + assert de presencia/estructura + guía documentada** (fail-fast si falta la hoja/bloque, ceros documentados solo si T0-0.4 ordena ceros) | Asumir editable o asumir "siempre ceros" antes de T0 | V2/V8: no hay fuente conocida en insumos. Inventar valores violaría la regla de hierro; ignorar la hoja dejaría la última hoja sin HU |
| D3 | L-menores doble desenlace T0: (a) operando real → mapa editable por Id; (b) cero estático manual → protegida + test que fija el 0 (gemela F178/F521 HU-07T0 §0.2) | Mapear por proximidad ("la L más cercana") o forzar 0 global | La columna L ya tiene semántica mixta (operando vs 0 explícito, V4); solo la evidencia celda-por-celda distingue |
| D4 | Si hay fuente de interventoría futura: finder `BuscarInterventoria` prefix-based sin fechas, match normalizado sin diacríticos (patrón HU-11 D4) | Prefijo con fechas o literal con acento | Rango heterogéneo (precedente ASE4 1607–3107) y encoding degradado en disco lo prohíben |
| D5 | Validación protegida **por período** para el bloque (Q1 = filas T0-Q1; Q2 = filas T0-Q2-canónico), matcheo estricto `Single` por Id | Invariante único de filas fijas | El Detalle dice "según período": filas fijas globales son hipótesis, no hecho |
| D6 | Misma pasada de escritura HU-07..HU-12 (overload lista; un `File.Copy` + validación pre/post con mapa HU-16) | Segundo pase / writer separado | La superficie nueva —si la hay— son celdas del mismo modo leaf. Un segundo pase rompería atomicidad y hash A4 |
| D7 | Golden HU-16 = **ambos canónicos reutilizados** (sin re-fijar) + Capa A (§2.7); ningún oráculo nuevo | Nuevo golden o "elegir en el test" | Dos oráculos por período = merge verde en falso; A5 (prohibido caché-salida vs golden) aplica |
| D8 | Modelos aditivos nullable (patrón `AjustesSfT = null`): sin insumos HU-16 = comportamiento HU-15 puro | Campos obligatorios en `WorkbookLeafInputs` | Compatibilidad hacia atrás por construcción; Q1/Q2 existentes no pagan lo que no usan |

### 2.3 Escritura HU-16 (misma sesión atómica)

```text
File.Copy canónico → salida (una vez, igual que HU-07..HU-15)
  └─► ValidarFormulasProtegidas (mapas HU-07..HU-12 intactos + mapa HU-16 T0:
      INTERVENTORIA visibles/editables según D2; L-menores según D3;
      resto 2.7 / DetValiRetri / VALIDACION_* siempre protegidas)
        └─► EscribirCeldasLeaf HU-07..HU-12 (intactas; solo activas según período)
        └─► EscribirCeldasInterventoria (solo editables T0-D2a; en D2b: nada)
        └─► EscribirCeldasLEspecialesMenores (solo operando T0-D3a; en D3b: nada)
              └─► revalidar fórmulas protegidas → guardar
```

Ante cualquier fallo: borrar salida parcial (patrón existente). Plantilla origen jamás mutada (hash A4 extendido al mapa HU-16). Sin capacidad de filas: fail-fast honesto (HU-07), nunca insert/delete.

### 2.4 Dominio (Core, sin deps)

- `InterventoriaAseInputs` (si D2a): un ASE — celdas refs T0 → valor (K/M/N según veredicto: ¿M+N=K formulado o tres valores externos?) + `TotalInterventoria` como propiedad calculada con la aritmética congelada por T0-0.3. Si D2b: **no se crea** (la hoja queda fuera del dominio, con assert estructural en el mapa protegido + guía en `CatalogoErrores` si falta).
- `LEspecialesMenoresAseInputs` (si D3a): por ASE — dict refs T0 → valor para las L-menores operando; `TieneCeldaLMenor`/valores con distinción "leído 0" vs "slot ausente" (doctrina HU-12: nunca 0 silencioso en operando).
- `WorkbookLeafInputs`: extensión mínima aditiva nullable (patrón `AjustesSfT = null`, D8).
- `ICalculoRemuneracion`: overloads aditivos solo si T0 demuestra que interventoría alimenta `TotalAse`/consolidado; si el bloque es informativo (no alimenta D104:D108), **no hay overload de cálculo** — solo lectura/validación/golden (cierre honesto, sin rebase).
- `ProcesadorPeriodo`: integra lectura HU-16 al flujo por ASE en el modo que T0 justifique, con fail-fast que nombra ASE + hoja + celda; sin insumos HU-16 el path se omite (mismos asserts que hoy).
- Coherencia `WorkbookLeafCoherence`: `ValidarContraResultado` itera HU-16 con matcheo estricto `Single` por Id (misma regla HU-07, extendida).
- `CatalogoErrores` (HU-14): sin códigos nuevos; los fail-fast HU-16 usan `ERR-FUENTE-NO-ENCONTRADA` / `ERR-FORMATO-FUENTE` / `ERR-VALIDACION` / `ERR-PLANTILLA` según fase (guías existentes + hoja/celda nombradas).

### 2.5 `IValidador` HU-16 (sin reabrir HU-04..HU-15)

1. Todo lo HU-07..HU-15 intacto (matcheo estricto por `Ase.Id`, `GranTotal = Σ`, gates leaf/empresa/banco/balance/ajustes/DetRetri/cruzadas, niveles HU-14).
2. Nuevo, **solo si D2a/D3a**: por cada ASE, valor de dominio == evidencia T0 ±0.5 (matcheo `Single` por Id; insumo ausente → error que nombra ASE+hoja+celda).
3. Nuevo, **en D2b/D3b**: assert estructural (bloque presente con el carácter T0; L-menor sigue 0/valor-fijo) — si diverge, fail-fast nombrado (detecta stale futuro: el riesgo que motiva esta HU).
4. Q1/Q2 existentes sin cambios: ningún gate nuevo activo fuera de su período/modo justificado (regresión 259/259 ciega).
5. El validador NO abre `.xlsx`. La composición del bloque va a Capa A (dominio vs caché golden), no a gates de fórmula.
6. Extensión 2.7 **solo si la hoja lo exige** (Req 5): con recorte declarado en §9; si no lo exige, 2.7 ni se menciona en el diff.

### 2.6 UI — Visual Design Intent (delta mínimo)

Densidad Balanced, mismos GroupBoxes, sin restyle/colores/iconos. El `txtLog` agrega, por cada ASE (en el modo que T0 justifique), una línea INTERVENTORIA (origen/valores por ASE, "esperado post-Excel" o "insumo externo declarado — hoja intacta" en D2b) + línea L-Especiales-menores si D3a. Serilog: mismos eventos HU-07..HU-15 con propiedad `Hoja = "INTERVENTORIA"`. Sin nuevos controles. La CLI imprime la misma línea final grepable (`RESULTADO … runId=…`); sin flags nuevos salvo exigencia justificada (G8).

### 2.7 Golden Capa A HU-16 (honestidad HU-06..HU-15)

| # | Qué | Contra qué | Tol |
|---|---|---|---|
| A1 | Celdas HU-16 escritas en la **salida** (solo editables T0-D2a/D3a; en D2b/D3b: vacío — el test lo declara) | Mismas celdas **leaf** de ambos canónicos | ±0.5 |
| A2 | Valores de **dominio** HU-16 por ASE (aritmética T0-0.3/0.5) | Caché golden del bloque + L-menores, ambos canónicos | ±0.5 |
| A3 | INTERVENTORIA no-escrita + L-menores estáticas + `VALIDACION_*` + resto protegido siguen siendo fórmula / intactas en la salida | Estructura | n/a |
| A4 | SHA256 de ambos canónicos igual antes/después | — | n/a |
| A5 | **Prohibido** comparar caché de fórmula de la salida vs golden; **prohibido** usar agregados HU-02 o totales existentes como oráculo HU-16 | — | prohibido |
| A6 | `TotalAse`/`GranTotal` de dominio (si D2a los toca) vs caché golden D104:D109, ambos canónicos | Caché golden | ±0.5 |
| A7 | Q1/Q2 intactos: suites 259/259 + harness 24/24 re-assert sin duplicar suite | Goldens existentes | ±0.5 |
| A8 | Stale-guard: L-menor declarada-0 que en algún canónico venga no-cero → el test lo revela (falla nombrando ASE+celda) | Ambos canónicos | n/a |

Capa B (manual Excel: abrir, recalcular, comparar bloque + D104:D109 vs canónicos) fuera de CI, protocolo §5.3. **Capa B manual queda como acción del usuario** (la instruye HU-17).

### 2.8 File changes previstos

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapInterventoria.cs` | Crear | Mapa explícito por (`Ase.Id`, período): visibles T0 + editables T0 + protegidas (D1/D5) |
| `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` | Modificar | Dispatch HU-16 por período/ASE (D2a/D3a); path existente intacto |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Escritura HU-16 en la misma pasada (o cero celdas en D2b/D3b) + mapa protegido por período |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs` | Modificar | Coherencia HU-16-vs-consolidado, matcheo estricto |
| `Remuneracion.Infrastructure/FileSystem/ArchivoFuenteLocator.cs` | Modificar **solo si D2a con fuente** | `BuscarInterventoria` agnóstico a rango/diacríticos (D4); si D2b: no se toca |
| `Remuneracion.Core/Interfaces/ILocalizadorArchivosAse.cs` | Modificar **solo si D2a con fuente** | Firma del finder nuevo (aditiva) |
| `Remuneracion.Core/Models/InterventoriaAseInputs.cs` (+ `LEspecialesMenores…` si D3a) | Crear **solo si D2a/D3a** | Superficie editable evidenciada (§2.4) |
| `Remuneracion.Core/Models/WorkbookLeafInputs.cs` | Modificar (aditivo nullable) | Extensión D8 (default = HU-15 puro) |
| `Remuneracion.Core/Interfaces/ICalculoRemuneracion.cs` | Modificar **solo si T0 demuestra alimentación al total** | Overloads aditivos; firmas viejas intactas |
| `Remuneracion.Core/Services/CalculoRemuneracion.cs` | Modificar **solo si aplica** | Implementación del overload HU-16 |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Modificar | Gates §2.5 (por período/modo; existente intacto) |
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` | Modificar | Integración HU-16 al flujo por ASE (fail-fast ASE+hoja+celda) |
| `Remuneracion.WinForms/Form1.cs` (+ `Remuneracion.Cli` si G8 lo exige) | Modificar | Resumen INTERVENTORIA por ASE (§2.6) + Serilog `Hoja` |
| `Remuneracion.IntegrationTests/GoldenInterventoriaTests.cs` | Crear | Capa A HU-16 (§2.7, ambos canónicos) |
| `Remuneracion.IntegrationTests/InterventoriaTests.cs` | Crear | Dominio: veredictos D2/D3, mismatch nombra ASE+hoja+celda, stale-guard A8 |
| `Remuneracion.IntegrationTests/ProcesadorPeriodoTests.cs` | Modificar | Casos HU-16 (integración período + negativas) |
| `Remuneracion.IntegrationTests/Insumos.cs` | Modificar | Helpers `Interventoria(aseId)`, `LEspecialesMenores(aseId)` + período/modo T0 |

**No tocar (salvo bug blocker):** `IPlantillaWriter`/validation-only (HU-04); agregados HU-02; coherencia `F25`-Extemp HU-05; mapas HU-07..HU-12; `DetRetriRounder`; reglas Q1/Q2 de `ValidadorBasico`; `CodigosSalida`/`CatalogoErrores` (sin códigos nuevos); cadena 2.5/2.6 certificada; recorte HU-11 (HU-08 fuera de Q2); `requirements/` legado; harness auxiliar.

---

## 3. SPEC / Acceptance Criteria (trazabilidad Propuesta §10 + base docs)

### Requirement 1 — T0 propio con veredicto que bloquea el mapa (CA-4; Detalle + HU-07T0)

El T0 **MUST** producir: (a) tabla valor-vs-fórmula celda por celda del bloque K25:N31 + filas 30–31 en ambos canónicos con `<f>`/`<v>`; (b) enumeración L-completa R1 por ASE en ambos canónicos con dif contra el set V4; (c) veredictos D2/D3 escritos. **MUST NOT** entrar ninguna celda HU-16 al código sin fila en estas tablas.

- GIVEN ambos canónicos dumpeados (unzip + `<c>` raw, método HU-07) → THEN tablas T0-0.2/0.3/0.5 congeladas con SHA256 de trabajo.
- GIVEN veredicto ausente o ambiguo → THEN NEEDS_CONTEXT con recorte (nunca invención).

### Requirement 2 — INTERVENTORIA según veredicto, nunca inventada (CA-1/CA-2/CA-3; Prompt Vo + Detalle)

El sistema **MUST** implementar exactamente el desenlace T0: D2a → leer/escribir/validar lo evidenciado en la misma pasada; D2b → hoja protegida intacta + assert estructural + guía documentada (fail-fast si falta el bloque, o ceros documentados si T0-0.4 los ordena). **MUST NOT** inventar ningún costo: sin fuente o sin cierre ±0.5 → NEEDS_CONTEXT con recorte.

- GIVEN D2a con aritmética T0 (p. ej. M+N=K o K externo fijo) → THEN dominio vs caché golden ±0.5 por ASE (A2) + A1 verde.
- GIVEN D2b (insumo externo sin fuente) → THEN cero escrituras HU-16, hoja intacta (A3), log declara "insumo externo — hoja intacta", test fija el estatuto.
- GIVEN bloque ausente o con `<f>` sin `<v>` en celda de gate → THEN fail-fast que nombra ASE + `INTERVENTORIA` + celda (doctrina S-4 HU-14).

### Requirement 3 — L-Especiales menores mapeadas o fijadas-en-0 (CA-1/CA-2; Detalle + follow-up HU-07)

Cada celda L numérica fuera del set V4 **MUST** tener veredicto D3: operando (→ mapa editable por Id, cierra contra fuente ±0.5) o cero estático (→ protegida + test que fija el 0 en ambos canónicos, gemela F178/F521). **MUST NOT** quedar ninguna L numérica sin veredicto.

- GIVEN L-menor con fuente que cierra → THEN D3a: lectura header/rango-driven + A1/A2 verdes.
- GIVEN L-menor 0 en ambos canónicos sin fuente → THEN D3b: no se escribe; A8 la vigila contra stale futuro.
- GIVEN L-menor no-cero en Q2 canónico sin fuente que cierre → THEN NEEDS_CONTEXT con recorte (nunca 0 silencioso en operando).

### Requirement 4 — Integración multi-ASE + Serilog + paridad CLI↔UI si aplica (CA-6/CA-7; §6)

El procesador **MUST** integrar HU-16 al flujo por ASE (fail-fast ASE+hoja+celda; una escritura) y el log **MUST** resumir INTERVENTORIA por ASE con `Hoja = "INTERVENTORIA"` + `RunId`. La paridad CLI↔UI **MUST** demostrarse: sin flags nuevos, igualdad de valores entre vías (patrón HU-15 §2.7); con flags nuevos justificados, matriz de parseo + paridad. **MUST NOT** cambiar valores/gates existentes.

- GIVEN ejecución 5 ASE con HU-16 → THEN mismos `Resultado` + snapshots-oráculo iguales entre vía CLI in-process y vía directa.
- GIVEN modo/fuente HU-16 ausente → THEN comportamiento HU-15 puro (extensión nullable, D8).

### Requirement 5 — Guardas fuera de alcance (CA-4)

`DetValiRetri`/`VALIDACION_*`, `ANT EXT-REV`, cadena 2.5/2.6, conciliación HU-08 en Q2 **MUST** seguir protegidas/intactas. La extensión 2.7 **MUST NOT** ocurrir salvo exigencia de la hoja con recorte declarado en §9. Si la plantilla no alcanza → fail-fast honesto, **MUST NOT** insert/delete de filas.

| CA §10 | HU-16 |
|---|---|
| CA-1 | Lee lo evidenciado T0 (fail-fast nombra ASE+hoja+celda; locator agnóstico solo si hay fuente) |
| CA-2 | Cada celda HU-16 = evidencia ±0.5 o estatuto declarado con test (A2/A8) |
| CA-3 | Capa A ambos canónicos (A1–A8); INTERVENTORIA correcta post-Excel (Capa B residual) |
| CA-4 | Reassert mapa por período + canónicos no mutados (hash); sin insert/delete |
| CA-5 | Gates nuevos por período/modo + coherencia con matcheo estricto; 2.7 solo si la hoja lo exige |
| CA-6 | Serilog + log INTERVENTORIA por ASE con `RunId` |
| CA-7 | Mismo flujo 5-ASE (+ CLI si aplica) con resumen HU-16 |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 500–1000 (depende de D2/D3: D2b+D3b ≈ 300–500; D2a/D3a ≈ 700–1000) |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 T0 doble + mapas congelados → PR2 dominio + cálculo (solo si D2a/D3a) → PR3 readers + locator (solo si hay fuente) → PR4 writer + validador + período + UI/CLI → PR5 golden ambos canónicos + regresión |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes. Chained PRs recommended: Yes. Chain strategy: pending.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 0 | Discovery T0 doble + mapas congelados + veredictos D2/D3 | PR 1 | Bloquea todo; solo lectura + datos |
| 1 | Modelos + cálculo (solo si D2a/D3a lo exigen) | PR 2 | Depende de PR 1; en D2b/D3b esta unidad es vacía declarada |
| 2 | Readers + locator (solo si hay fuente evidenciada) | PR 3 | Depende de PR 1; en D2b puede no existir |
| 3 | Writer + validador + período + UI/CLI | PR 4 | Depende de PR 1 |
| 4 | Golden ambos canónicos + paridad + regresión 259/259 + 24/24 | PR 5 | Depende de PR 2–4 |

### Phase 0 — Discovery T0 (BLOQUEANTE, solo lectura, réplica del método HU-07/T0)

- [x] 0.1 Fijar SHAs de trabajo de ambos canónicos (Q1 + `Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx`); registrar que no se re-fijan. La plantilla Q2-raíz queda como control, nunca oráculo. **Acta W-1 (completada en HU-17):** ver bloque "Acta W-1 — SHA256 de trabajo" al final de §4 y el instructivo Capa B (HU-17).
- [x] 0.2 Dumpear por OpenXML (método HU-07: unzip + `<c>` raw) el bloque INTERVENTORIA K24:N32 en ambos canónicos y clasificar celda por celda: valor editable vs fórmula (+ `<f>` y `<v>` de cada una). Registrar si M/N son fórmula de K, valores externos, o referencias a otra hoja. (Veredicto: K/M/N = valores estáticos; K31/M31/N31/K32 = SUM.)
- [x] 0.3 Dictaminar el **origen del costo**: ¿alguna celda referencia otra hoja/fuente, o es insumo externo puro? Buscar exhaustivamente `*nterventoria*` en `Docs/Insumos/` (ambos períodos, normalizado sin diacríticos). Registrar veredicto D2a/D2b por período. (Veredicto: D2b — sin fuente.)
- [x] 0.4 En D2b: dictaminar el estatuto honesto — ¿hoja intacta con assert estructural (default), o ceros documentados? ¿fail-fast si falta el bloque? Registrar la guía de usuario (deuda de documentación para HU-17). (Estatuto: hoja protegida intacta + assert estructural + fail-fast; guía → Manual HU-17 §9.)
- [x] 0.5 Enumerar la columna L **completa** de `Reporte Componentes R1` por ASE en ambos canónicos (toda celda L numérica o con fórmula) y diferir contra el set V4; clasificar cada diferencia: operando candidato (D3a) vs cero/valor-fijo estático (D3b). **Nada L-menor entra al código sin esta tabla.** (D3a: 113 celdas no-cero mapeadas por rol/ocurrencia; D3b: estáticas.)
- [x] 0.6 Para cada L-menor candidata D3a: verificar cierre contra fuentes `REMUNERACION 202607{1,2}/{1..5}-*/` ±0.5; si no cierra → veredicto fallido parcial (NEEDS_CONTEXT), no invención.
- [x] 0.7 Congelar `WorkbookLeafCellMapInterventoria` (+ extensión L-menores): visibles + editables + protegidas (incl. filas 30–31, resto 2.7, DetValiRetri/VALIDACION_*, ANT EXT-REV). Check de capacidad de filas (si no alcanza: fail-fast honesto, sin insert/delete).
- [x] 0.8 Confirmar baseline: Q1/Q2 existentes intactos (259/259 + 24/24 en verde antes de empezar; si algo falla → NEEDS_CONTEXT, no se empieza).

### Phase 1 — Dominio (solo lo que T0 justifique)

- [x] 1.1 `InterventoriaAseInputs.cs` (D2a) / `LEspecialesMenoresAseInputs.cs` (D3a) con propiedades calculadas de aritmética T0 + `WorkbookLeafInputs` aditivo nullable (D8). En D2b/D3b: esta fase se declara vacía con evidencia. (D2b: sin modelo Interventoria; D3a: `LEspecialesMenoresAseInputs` creado.)
- [x] 1.2 Overloads `ICalculoRemuneracion` + implementación **solo si T0 demuestra alimentación al total** (§2.4); si el bloque es informativo, cierre honesto sin overloads. (Informativo: sin overloads.)
- [x] 1.3 Tests in-memory de cálculo/veredicto: D2a/D3a cuadran ±0.5; D2b/D3b fijan estatuto; origen-equivocado (p. ej. agregado HU-02 como interventoría) falla a propósito.

### Phase 2 — Lectura (solo si hay fuente evidenciada)

- [x] 2.1 `BuscarInterventoria` (D4: sin fechas, match sin diacríticos) + firma en `ILocalizadorArchivosAse` — **solo si D2a con fuente**; si D2b no existe esta tarea. **(N/A — D2b: sin fuente; cero cambios de locator.)**
- [x] 2.2 Readers header/rango-driven HU-16 (Especiales-opcional donde aplique; fallo nombra ASE+hoja+celda; distinguir "leído 0" de "slot ausente"). (`ExcelDataReaderWorkbookLeafInputReader.LeerLEspecialesMenores`.)
- [x] 2.3 `Insumos.cs`: helpers `Interventoria(aseId)`, `LEspecialesMenores(aseId)` + período/modo T0. **(Resuelto sin helpers nuevos: L-Especiales se leen de la ruta R1 existente `Insumos.R1(aseId)`/`R1Q2(aseId)`; D2b no requiere ruta.)**

### Phase 3 — Orquestación + UI/CLI delta mínimo

- [x] 3.1 `OpenXmlPlantillaWriter`: mapa protegido por período (D5) + escritura de editables T0 (o cero celdas en D2b/D3b) en la misma pasada (D6). (D2b: cero celdas INTERVENTORIA + assert estructural; D3a: L-menores escritas.)
- [x] 3.2 `ValidadorBasico` + `WorkbookLeafCoherence`: gates §2.5 (por período/modo + matcheo estricto; existente intacto; 2.7 solo si Req 5).
- [x] 3.3 `ProcesadorPeriodo`: integración HU-16 al flujo por ASE (fail-fast ASE+hoja+celda); sin insumos HU-16 → omite (D8).
- [x] 3.4 `Form1` (+ CLI si G8): resumen INTERVENTORIA por ASE (§2.6) + Serilog `Hoja = "INTERVENTORIA"`; `OpcionesCli` solo si hay flags nuevos justificados. (Sin flags nuevos.)

### Phase 4 — Pruebas y evidencia

- [x] 4.1 `InterventoriaTests` (in-memory + fuentes reales sin salida): veredictos D2/D3, Especiales-ausente→0 donde aplique, header ausente → fallo ASE+hoja+celda, mismatch nombra, origen-equivocado prohibido.
- [x] 4.2 `GoldenInterventoriaTests`: matriz §2.7 contra **ambos canónicos** (A1–A8; A8 stale-guard incluido).
- [x] 4.3 Paridad CLI↔UI (Req 4, G8): igualdad de valores entre vías + `runId` correlacionado; si hay flags nuevos, matriz de parseo. (Cubierto por la suite paridad existente; sin flags nuevos.)
- [x] 4.4 Negativas: slot T0 ausente (falla ASE+hoja+celda); `<f>` sin `<v>` en celda de gate (falla nombrando, doctrina S-4); salida == plantilla (no in-place); plantilla sin filas (fail-fast, sin insert/delete); rama HU-16 con path viejo (bloqueado, intacto).
- [x] 4.5 Los 259 tests + harness 24/24 existentes verdes; build 0 warnings; CRLF; sin commit. (264/264 + 24/24 al cierre HU-16; re-verificado HU-17 U0.)

### Phase 5 — Documental

- [x] 5.1 Capa B manual §5.3 ejecutada una vez sobre ambos canónicos y evidenciada (acción del usuario; sin fingirla como gate de merge). **(Acción del Ingeniero, pendiente: se ejecutará con el instructivo HU-17 — `Docs/Instructivo-Capa-B.md`.)**
- [x] 5.2 Cierre deja explícito: estatuto D2b (si aplica) como deuda documentada para el manual HU-17 (origen del costo, cómo diligenciarlo), y el frente HU-17 (siguiente HU). (Estatuto D2b documentado en Manual HU-17 §9 + Instructivo §6.)

---

### Acta W-1 — SHA256 de trabajo (completada en HU-17)

Fijados como ancla de integridad (HU-17 §0.1 V2; no se re-fijan). La copia canónica vive en `Docs/Instructivo-Capa-B.md` §2 y `Docs/Manual-Usuario-Remuneracion-UAESP.md` §12.

| Archivo | Rol | SHA256 |
|---|---|---|
| `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` | Golden Q1 | `0B090E9C4851D86DAC534F2965E4A38393CC5BC09350C2A824FD48F7931C096F` |
| `Docs/Insumos/Remuneracion 202607-2 Total.xlsx` | Golden Q2 (valores) | `584310105AC7223CE840833E3CB64E26F95B61907C9ECD959A0A632EEC21DCB1` |
| `Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx` | Canónico Q2 | `95825422B32FBE9E2B60F35C9638E492455FB98CFA182976B657FAC3520E0B8C` |
| `Docs/Insumos/REMUNERACION 2026072/Plantilla  _ Remuneracion 202607-2 Total.xlsx` | Control Q2 (nunca oráculo) | `509BF210435138B3D36947582E295F27CE1F4D9452BEB1A8D752F3EF4F57529B` |

---

## 5. Test Strategy

| Capa | Qué | Cómo |
|---|---|---|
| Unidad | Veredictos D2/D3 (aritmética T0, estatuto declarado, `Single` por Id) | In-memory |
| Unidad | Cálculo HU-16 si D2a toca totales (cuadra; path viejo intacto) | In-memory |
| Unidad | Readers HU-16 si hay fuente (headers/rangos, Especiales opcional, 0 vs ausente) | Fuentes reales, sin Excel de salida |
| Integración | Período con HU-16 (carpetas reales, salida temp, fail-fast ASE+hoja+celda) | Insumos ambos períodos, canónicos |
| Golden Capa A | Matriz §2.7 ambos canónicos (A1–A8, incl. stale-guard) | OpenXML read-only + aritmética dominio (A5 aplica) |
| Regresión | 259/259 + harness 24/24 verdes | Suites existentes, sin cambios |
| UI/CLI | Resumen INTERVENTORIA + paridad de valores | Funcional manual + paridad in-process (sin harness) |
| Capa B | Bloque + D104:D109 post-Excel, ambos canónicos | Manual — §5.3 |

### 5.1 Fixtures

- Goldens/plantillas: ambos canónicos (sin re-fijar; Q2-raíz = control).
- Fuentes: `Docs/Insumos/REMUNERACION 2026071/{1..5}-*/` + `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/` (+ búsqueda `*nterventoria*` normalizada T0-0.3).
- Regresión: goldens + fuentes existentes (intactos).
- Referencia: tablas T0-0.2/0.5 + set V4, tolerancia ±0.5.

### 5.2 Casos negativos obligatorios (nombran ASE + hoja/celda)

Bloque INTERVENTORIA ausente (fail-fast ASE+hoja); `<f>` sin `<v>` en celda de gate (falla nombrando); L-menor no-cero sin fuente que cierre (NEEDS_CONTEXT, nunca 0 silencioso); origen-equivocado como interventoría (prohibido — el test lo demuestra fallando); salida == plantilla (no in-place); plantilla sin filas (fail-fast, sin insert/delete); rama HU-16 con path viejo (bloqueado, intacto); segundo oráculo Q2-raíz como oráculo (prohibido).

### 5.3 Protocolo manual Capa B (no CI — acción del usuario)

1. Generar salida de cada período a ruta distinta del canónico. 2. Abrir en Excel, recalcular. 3. Comparar bloque INTERVENTORIA + L-menores + D104:D108 + D109 vs canónicos ±0.5. 4. Evidencia adjunta. No es gate de merge.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | HU-16 = extensión de lectura/escritura/validación (mapa hermano + ramas); cálculo/período/writer conservan su rol HU-07..HU-15. En D2b/D3b la extensión es solo aserto estructural, sin dominio nuevo. |
| **O** | Se agregan mapa/modelos/overloads solo donde T0 los justifica; los paths existentes funcionan con la extensión en `null` (abierto sin modificar). |
| **L** | El reader leaf suma rama HU-16 sin cambiar el comportamiento certificado ni los fail-fast existentes. |
| **I** | `InterventoriaAseInputs`/`LEspecialesMenores…` separados de `BalanceScInputs`/`AjustesSfTInputs`; reader/validador/locator crecen por overload. Sin códigos de error nuevos. |
| **D** | Core define inputs/gates; Infrastructure/WinForms/Cli componen. Sin nuevas deps ni paquetes. |

### 6.2 Best Practices

- La verdad del workbook manda: valor-vs-fórmula celda por celda en ambos canónicos antes de codificar (Rector §6 preservación; método HU-07).
- El origen del costo no se inventa: sin fuente o sin cierre ±0.5 → NEEDS_CONTEXT con recorte (regla de hierro §0.1).
- Mapa explícito por `Ase.Id` + período verificado, no offsets ni filas/columnas fijas (Rector §11.3; el Detalle advierte "según período").
- Cero estático con test que lo fija (gemela F178/F521), nunca 0 silencioso en operando (doctrina S-4 HU-14).
- Una escritura atómica; canónicos nunca mutados; hash A4 extendido.
- Golden honesto ambos canónicos con stale-guard A8 (OpenXML no recalcula; A5).
- Fail-fast nombra ASE+hoja+celda; sin salida certificada ante fallo; sin insert/delete de filas.
- Q1/Q2 blindados por construcción (dispatch + extensión nullable + 259/259 + 24/24 como red).
- Paridad CLI↔UI mínima y demostrable (G8); sin flags ni códigos nuevos salvo exigencia justificada.

### 6.3 Performance

- 5 ASE × (bloque INTERVENTORIA + barrido L-menor) en la misma sesión OpenXML + hasta 5 resoluciones por prefijo (solo si hay fuente). Irrelevante a esta escala; `Task.Run` existente para no congelar el form; CLI sincrónica como hoy.

**Veredicto:** APROBADO como cierre de la última hoja sin HU **si** T0 congela ambos mapas con evidencia (§4 Fase 0, veredictos D2/D3 escritos) y se acepta CA-3 parcial (Capa A en CI, Capa B manual como acción del usuario instruida en HU-17).

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| El bloque es insumo externo puro sin fuente (V2/V8 se confirma) y el "cierre" es solo declaración | Alta | Es el desenlace D2b diseñado: protegida + assert + guía para HU-17; el plan lo contempla sin rebase; NEEDS_CONTEXT solo si ni el estatuto es dictaminable |
| M/N resultan fórmulas de K (mitades) y K es el único externo | Media | D2a parcial: K declarado-externo + M/N validadas como fórmula; T0-0.2 lo distingue antes de codificar |
| El bloque se desplaza "según período" y Q2 canónico trae otro rango | Media | T0-0.2 en ambos canónicos + validación protegida por período (D5); si diverge sin patrón, recorte al período verificado |
| L-menor no-cero en Q2 canónico sin fuente que cierre | Media | T0-0.6; veredicto fallido parcial + NEEDS_CONTEXT con recorte, nunca 0 silencioso |
| Ceros legítimos confundidos con "falta de lectura" | Alta | El reader distingue "leído 0" de "slot ausente" (fallo); tests de ceros explícitos; A8 stale-guard |
| Q2-raíz usada como oráculo por accidente (valores iguales a Q1 la hacen tentadora) | Media | T0-0.1 la declara control por escrito; A5 + tests que fallan si se usa como oráculo |
| Inflar a HU-17 (manual/paquete) o a 2.7 (DetValiRetri tienta) dentro de esta HU | Media | §0.2 out-of-scope + Req 5; rechazar PRs que lo metan |
| Comparar caché de salida vs golden y "cerrar" en falso | Alta | A5 + canónicos únicos; este plan lo prohíbe |
| Q1/Q2 regresión rota por la extensión | Media | 259/259 + 24/24 como red en PR5; extensión nullable + dispatch (G4/D8) |
| Reabrir cadena 2.5/2.6 o HU-08-en-Q2 por inercia | Media | Fuera de alcance explícito; rechazar PRs que lo metan |

---

## 8. Rollback

- Revertir PRs en orden inverso (golden/paridad → UI/CLI/validador → writer → readers/locator → dominio → mapas/T0).
- HU-01..HU-15 intactas sin esta HU: extensión en `null` = comportamiento HU-15 puro; `INTERVENTORIA` sigue protegida (baseline V6); L-menores siguen como hoy.
- `Docs/Insumos/` untracked: jamás como destino; pisada se restaura desde backup.
- Si T0 demuestra bloque sin fuente y sin estatuto dictaminable, o L-menor que exige insert/delete (Riesgos 1–4), recorte con rebase a lo verificado, no invención.

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §9 Fase 3 / §6 / §10 (±0.5) + Detalle (INTERVENTORIA L25:N31 por ASE según período; columna "Especiales") y Prompt Vo (costo de interventoría por ASE) como origen funcional (§1.5). Manual/entrega (HU-17), cambios de cálculo/escritura existentes, 2.7 (salvo exigencia de la hoja), DI y restyle quedan fuera. Proceso de Recaudo NO aplica.
2. T0 doble bloqueante (G1): T0-INTERVENTORIA (valor-vs-fórmula + origen, ambos canónicos) y T0-L-ESPECIALES (columna L completa vs set V4). Ninguna celda HU-16 entra al código sin evidencia; sin fuente o sin cierre ±0.5 → NEEDS_CONTEXT con recorte, nunca invención.
3. INTERVENTORIA doble desenlace sin rebase (D2): editable-evidenciada (D2a) vs insumo-externo-declarado con fail-fast o ceros documentados (D2b).
4. L-menores doble desenlace sin rebase (D3): operando real (D3a) vs cero estático fijado con test, gemela F178/F521 (D3b); A8 stale-guard contra futuro.
5. Mapa hermano por (`Ase.Id`, período) con validación protegida por período (D1/D5); prohibidos offsets, `if`-parches e insert/delete.
6. Integración aditiva al procesador multi-ASE + resumen Serilog (`Hoja = "INTERVENTORIA"`) + paridad CLI↔UI mínima (G8); sin códigos/flags nuevos salvo exigencia justificada.
7. Golden Capa A ambos canónicos con honestidad HU-06..HU-15 (A1–A8); Capa B manual residual como acción del usuario (la instruye HU-17).
8. Q1/Q2 intactos por construcción (dispatch + extensión nullable + 259/259 + 24/24 ciegos); build 0 warnings; CRLF; sin commits.
9. UI delta mínimo; OPA = Ejecutar.
10. Apply espera aprobación + PRs encadenados con T0 al frente (Unidad 0–4, §4).

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
