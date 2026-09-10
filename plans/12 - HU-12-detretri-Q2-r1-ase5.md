# Plan 12 — HU-12: Cell-map Q2 completo por ASE/período + DetRetri Q2 (Fase 2, it. 2.6 ampliada) — RE-PLAN v2

> **Historia:** certificar el procesador Q2 end-to-end 5/5 replicando el método HU-07/T0 sobre la plantilla Q2: cell-map R1/R2/R4-Q2 completo por (`Ase.Id`, período) + DetRetri Q2 en enteros en la misma pasada, con Q1 bit-a-bit intacta por construcción (dispatch por período).
> **Rector (IRRENUNCIABLE):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 2 it. **2.6 DetRetri**, §6 arquitectura (capas, preservación de fórmulas, Serilog, lectura directa de fuente), §10 CA + tolerancia ±0.5. **Entra 2.6 ampliada al cell-map Q2 que la bloquea.** Salen 2.7 lógica, INTERVENTORIA, ANT EXT-REV — son HU posterior.
> **Origen funcional:** `Detalle de plantilla.docx` y `Prompt Maestro Vo.docx` en lo que digan de DetRetri/DetValiRetri (enteros redondeados) y del pegado R1/R2/R4 por empresa. `Proceso de Recaudo.docx` **NO aplica** salvo cita directa que lo exija.
> **Continuidad:** HU-01..HU-11 cerradas (ruta leaf multi-ASE atómica single-write, mapas explícitos por `Ase.Id`, validador estricto, UI modo 5-ASE, Golden Capa A, build 0/0, tests 105/105 —78 Q1 + 27 HU-11— + harness 24/24). **Cadena 2.5 ya certificada: NO se reabre.** Este plan NO reabre su semántica.
> **Golden Q2 canónico (no se re-fija):** `Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx`. Fuentes Q2: `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/`.
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-09 (re-plan v2; v1 queda conservada como §10 lecciones)

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. La evidencia V0 (abajo) se toma como dada —el orquestador la confirmó con dump propio de `sheet18.xml` del canónico— y **todo lo no cubierto por V0 lo dicta el T0 del re-plan, nunca se asume**. El apply espera aprobación explícita.

### 0.1 Evidencia V0 (verificada por el orquestador — se usa, no se re-descubre)

| # | Hecho | Consecuencia |
|---|---|---|
| V0.1 | CONSOLIDADO Q2: D9:D13 → `R1!F53/F206/F343/F468/F558` (**NO** F46/F176/F316/F437/F519); D47:D51 EXTEMP → `R1!F55/F208/.../F560`; D66 → `R4!D73` (**NO** D67); D85:D89 → `AJUSTES - SF-T!D47:D51` (igual que Q1 en forma) | El mapa HU-07/Q1 es **inválido en Q2**: ni los visibles ni los slots se reutilizan. Réplica completa del método HU-07/T0 sobre Q2, no parche |
| V0.2 | En el template Q2, F46/F176/F316/F437/F519/F521 son VALORES 0 o inexistentes; F478/F498 son FÓRMULAS de sub-bloques por empresa | La validación protegida HU-07 (`F46 debería seguir siendo fórmula`) es **falsa en Q2**: de ahí el fallo empírico del writer. Validación protegida parametrizada por período (D5) |
| V0.3 | ASE5-Q2: TOT_OPT = F558 (`F552+F531−L531`); fuente real 2 filas Mes/Total (filas 33 y 54): F33+F54−L33 = 13084233937.44 + (−1036320522.40) − 14901729.33 = 12033011685.71 = D13 golden | Variante R1-Q2-ASE5 cerrada aritméticamente (slots F531/F552/L531). EXTEMP ASE5 (F538/F512/L512 candidatos) **a confirmar por T0**, no asumido |
| V0.4 | DetRetri-Q2 **PROBADO**: `DetRetri_D = ROUND(CONSOLIDADO D104:D108, 0)` vía `DetRetriRounder` cierra ±0.5 en 5/5; hojas `DetRetri2026072` (sheet 93) y `DetValiRetri2026072` (sheet 94) existen; D9:D13 del canónico = valores 0 editables | Composición DetRetri-Q2 congelada desde el día uno: origen = D104:D108 por ASE. T0 solo verifica hojas/celdas destino, no la composición |
| V0.5 | Writer actual muere empíricamente contra el template Q2: `CalculoInvalidoException: La celda 'F46' de 'Reporte Componentes R1' debería seguir siendo fórmula` | El path Q2 muere **antes** en R1. M1 (`Replace 2026071→2026072` sí matchea hojas reales) queda pendiente de ejercitarse en CI tras el desbloqueo |
| V0.6 | Recorte 1 de HU-11 SE MANTIENE: conciliación HU-08 no se reescribe en Q2 (R4 ASE2-Q2 diverge) | Fuera de alcance explícito; ni T0 ni el apply lo tocan |
| V0.7 | `DetRetriRounder.Round` existe (`Remuneracion.Core/Rules/DetRetriRounder.cs`, ROUND/AwayFromZero, puro) | Reutilizado; prohibido duplicar la regla |

### 0.2 NO verificado — lo dicta el T0 del re-plan (§4 Fase 0), nunca se asume

1. R2-Q2: refs `E43…` por ASE (visibles + operandos editables exactos valor-vs-fórmula).
2. R4-Q2: refs `D73…` por ASE (visibles + operandos).
3. R1-Q2: visibles/operandos por ASE para ASE1–4 (solo ASE5 cerrado por V0.3) + EXTEMP Q2 por ASE (incluidos F538/F512/L512 candidatos de ASE5).
4. Hojas DetRetri/DetValiRetri Q2: columna D por ASE celda-por-celda (valor-vs-fórmula), capacidad de filas.
5. `ProtegidasBceParaPeriodo(true)` celda-por-celda contra el canónico (M1 a nivel estructura).

> **Regla de hierro del re-plan:** ninguna celda Q2 entra al código sin pasar por T0, salvo la composición DetRetri V0.4 y la aritmética ASE5 V0.3 (ya probadas contra el golden). Si algo no cierra ±0.5: NEEDS_CONTEXT con recorte, nunca invención.

### 0.3 Mapeo al Rector (in vs out)

**Entra porque §9 it. 2.6 / §6 / §10 lo piden ahora:**

| Rector | Qué cubre HU-12 v2 |
|---|---|
| §9 it. 2.6 | DetRetri Q2 como objetivo de escritura (enteros vía `DetRetriRounder`, composición V0.4) |
| §5 / §7.1 | Cell-map R1/R2/R4-Q2 completo por (`Ase.Id`, período) — réplica HU-07/T0 sobre Q2, incluida variante R1-Q2-ASE5 |
| §10 CA-1/CA-2 | Leer R1/R2/R4-Q2 de los 5 ASE + coherencia DetRetri = dominio ±0.5 por ASE |
| §10 CA-3/CA-4 | Writer multi-ASE Q2 certificable 5/5 + fórmulas intactas (mapa Q2 + M1 ejercitado) |
| §10 CA-5/CA-6/CA-7 | Validación Q2 por ASE, Serilog por ASE, UI sin cambios salvo resumen |

**Sale porque §9 lo asigna a 2.7 / Fase 3 (HUs posteriores) — EXPLÍCITO:**

- 2.7 lógica (`VALIDACION_TOTAL`, `VALIDACION_RECIP/ENEL/…`, `DetValiRetri` como lógica). `DetValiRetri`/`VALIDACION_*` Q2 entran al mapa **protegido**, no se implementan.
- `INTERVENTORIA` (L25:N31 por ASE — protegida), `ANT EXT-REV` (protegida).
- **Insert/delete de filas:** prohibidos. Si la plantilla no alcanza, **fail-fast honesto como en HU-07**.
- Reescritura de conciliación HU-08 en Q2 (recorte 1 HU-11, V0.6); restyle UI; DI framework.

### 0.4 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **T0 bloqueante con protocolo réplica-HU-07 sobre Q2** (§2.2/§4 Fase 0): dumps valor-vs-fórmula de cada visible y operando Q2 por ASE, mapas explícitos por Id+período, prohibidos offsets. V0.3/V0.4 se reutilizan, no se re-descubren. |
| G2 | **Cell-map Q2 hermano del mapa HU-07** (D1): el mapa Q1 queda intacto; Q2 vive en su propio mapa por (`Ase.Id`, `NumeroQuincena`). Dispatch por período en reader/writer/validación. |
| G3 | **DetRetri-Q2 con composición congelada V0.4**: `DetRetri_D(ase) = DetRetriRounder.Round(D104:D108(ase))`, escritura en la misma pasada; `DetValiRetri`/`VALIDACION_*` Q2 = protegidas. |
| G4 | **Validación protegida parametrizada por período** (D5): Q1 = fórmulas HU-07 intactas; Q2 = nuevo mapa T0. El invariante único actual queda derogado por falso en Q2 (V0.5). |
| G5 | **Q1 intacta por construcción:** variante Q2 solo activa con `NumeroQuincena == 2`; 105/105 + harness 24/24 como red de regresión ciega a Q2. |
| G6 | **Una sola llamada de escritura** por proceso (mismo overload lista HU-07..HU-11). Sin segundo pase. |
| G7 | **Golden Capa A Q2 completa incl. A6** (TotalAse/GranTotal Q2 cierra el pendiente HU-11); Capa B residual como acción del usuario. |
| G8 | **M1 se cierra aquí:** test obligatorio que ejercita el path Q2 del writer + `ProtegidasBceParaPeriodo(true)` contra hojas reales del canónico (sheets 93/94 incluidas). |

---

## 1. PROPOSE

### 1.1 Intent

Replicar el método HU-07/T0 sobre la plantilla Q2 para certificar el procesador Q2 5/5: cell-map R1/R2/R4-Q2 completo por ASE/período (con variante R1-Q2-ASE5) + DetRetri Q2 en enteros — todo en la misma pasada atómica, con Q1 bit-a-bit intacta.

### 1.2 In Scope

- Discovery T0 réplica-HU-07 sobre Q2 (§4 Fase 0: V0.2 pendiente) + `WorkbookLeafCellMapQ2` congelado solo con evidencia.
- Dispatch por (`Ase.Id`, `NumeroQuincena`) en reader leaf R1/R2/R4 (incluida variante 2-filas ASE5-Q2 con slots V0.3).
- Escritura DetRetri Q2 (enteros vía `DetRetriRounder`, composición V0.4) en la misma pasada + `DetValiRetri`/`VALIDACION_*`/`INTERVENTORIA`/`ANT EXT-REV` Q2 como protegidas.
- Validación pre/post parametrizada por período (Q1 = fórmulas HU-07; Q2 = mapa T0) + M1 obligatorio.
- Procesador Q2 end-to-end 5/5 + writer multi-ASE Q2 certificable + matriz Capa A completa (incluido A6).
- Resumen DetRetri por ASE en log/Serilog (delta mínimo UI).

### 1.3 Out of Scope

Todo §0.3 (2.7 lógica, INTERVENTORIA, ANT EXT-REV, insert/delete). Además: reescritura HU-08 en Q2 (V0.6); cadena 2.5 (certificada, no se reabre); motor Excel/COM en CI; restyle UI; framework DI.

### 1.4 Resultado de negocio

El Ingeniero ejecuta el modo 5 ASE sobre `REMUNERACION 2026072` con el golden canónico; obtiene `Remuneracion 202607-2 Total.xlsx` con R1/R2/R4-Q2 por ASE y DetRetri Q2 en enteros correctos post-Excel, D104:D109 actualizados; Q1 sigue bit-a-bit igual (105/105 + 24/24); el log audita DetRetri por ASE; las pruebas demuestran cell-map Q2 + DetRetri = dominio ±0.5 por ASE y procesador Q2 5/5.

### 1.5 Base documental (origen funcional — citas por sección)

| Documento | Sección / instrucción | Qué aporta a 2.6 ampliada |
|---|---|---|
| `Detalle de plantilla.docx` | Inst. hojas `DetRetri` / `DetValiRetri` (enteros como objetivo) + inst. de pegado R1/R2/R4 por empresa | Layout objetivo y semántica de cada tramo; base del mapa T0 |
| `Prompt Maestro Vo.docx` | Paso de pegado por empresa en valores + destino DetRetri/CONSOLIDADO | Qué tramos se pegan en valores (editables candidatas) vs calculados |
| Rector Propuesta | §9 it. 2.6 (alcance), §6 (preservación de fórmulas, lectura directa, trazabilidad), §10 CA + tolerancia ±0.5 | Rector normativo |
| `requirements/Fase1-Requerimientos.md` | DetRetri columna D = ROUND(valor,0), enteros | Composición de dominio (origen confirmado V0.4) |
| `plans/07 - HU-07 T0 Evidencia.md` | Protocolo 0.1–0.5 (dumps valor-vs-fórmula, mapa editable por ASE, decisión F178/F521, evidencia ±0.5) | **Patrón a replicar sobre Q2** — el T0 de este plan es su espejo Q2 |
| `Proceso de Recaudo.docx` | — | **NO aplica** salvo cita directa que lo exija |
| `README.md` / `AGENTS.md` | Redondear a entero antes de escribir (`DetRetri`) | Regla vigente, implementada por `DetRetriRounder` |

---

## 2. DESIGN

### 2.1 Tablas verificadas V0 (contratables desde el día uno)

**CONSOLIDADO Q2 (dump `sheet18.xml` del canónico — evidencia del orquestador):**

| Fila CONSOLIDADO | ASE1 | ASE2 | ASE3 | ASE4 | ASE5 |
|---|---|---|---|---|---|
| TOT_OPT D9:D13 | R1!F53 | R1!F206 | R1!F343 | R1!F468 | R1!F558 |
| EXTEMP D47:D51 | R1!F55 | R1!F208 | R1!… | R1!… | R1!F560 |
| R4 D66:D70 | R4!D73 | R4!… | … | … | … |
| AJUSTES D85:D89 | `AJUSTES - SF-T!D47:D51` (igual forma que Q1) | — | — | — | — |

**ASE5-Q2 cerrado (V0.3):** TOT_OPT F558 = F552+F531−L531; fuente 2 filas Mes/Total (filas 33 y 54): F33+F54−L33 = 13084233937.44 + (−1036320522.40) − 14901729.33 = 12033011685.71 = D13 golden. EXTEMP (F538/F512/L512 candidatos) → confirma T0.

**DetRetri-Q2 cerrado (V0.4):** `DetRetri_D(ase) = ROUND(D104:D108(ase), 0)` cierra ±0.5 en 5/5; hojas `DetRetri2026072` (sheet 93) + `DetValiRetri2026072` (sheet 94); D9:D13 canónico = valores 0 editables.

**Contratos y piezas que NO cambian:**

| Elemento | Estado | Tratamiento v2 |
|---|---|---|
| `DetRetriRounder.Round` (ROUND/AwayFromZero, puro) | Existe (V0.7) | Reutilizado; prohibido duplicar |
| `ConsolidadoAse.TotalAse` / `GranTotal` (getters, incluyen `AjustesSfT`) | Existen | Sin cambio; DetRetri deriva de D104:D108, no los altera |
| Path Q1 del reader leaf + mapa HU-07/Q1 (78 Q1) | Certificado | Intacto; Q2 vive en mapa hermano + dispatch por período |
| `ProtegidasBceParaPeriodo(false)` (Q1) | Certificado | Intacto; Q2 se verifica de verdad (M1) |
| Cadena 2.5 certificada (HU-11) | Certificada | NO se reabre |
| Recorte 1 HU-11 (HU-08 fuera de Q2) | Fijado (V0.6) | Se mantiene |

> Tablas R1/R2/R4-Q2 por ASE (visibles + operandos editables + EXTEMP por ASE) las extrae T0 del canónico + fuentes; este plan no las inventa salvo V0.3/V0.4.

### 2.2 Decisiones de arquitectura

| ID | Opción elegida | Descartada | Por qué |
|---|---|---|---|
| D1 | `WorkbookLeafCellMapQ2`: mapa hermano explícito por (`Ase.Id`, `NumeroQuincena`) — visibles Q2 (V0.1 + T0) + editables T0 + DetRetri-D (V0.4) | Parchear `WorkbookLeafCellMapPorAse` con `if Q2` o offsets aritméticos | OCP + Q1 intacta: Q2 tiene visibles distintos (F53≠F46, D73≠D67); mezclarlos rompería 78 Q1. R1-Q2 heterogéneo prohíbe offsets (doctrina HU-07 D1) |
| D2 | Variante leaf R1-Q2-ASE5 con slots V0.3 (F531/F552/L531; EXTEMP según T0): `LeerLeafInputs` despacha a rama ASE5-Q2; resto Q2 según mapa T0 | Flag genérico "2 filas" | El cambio es aditivo y localizado; el fail-fast de 3 filas sigue protegiendo Q1 y el resto de Q2 |
| D3 | DetRetri-Q2 con composición congelada V0.4 (sin doble desenlace: ya probada 5/5) | Re-descubrir la composición en T0 | V0.4 cierra ±0.5 en 5/5; T0 solo fija hojas/celdas destino. No se paga de nuevo lo ya probado |
| D4 | `WorkbookLeafCellMapDetRetriQ2` o sección DetRetri dentro del mapa Q2: dict explícito por `Ase.Id` (DetRetri-D esperado + refs destino T0) | Reutilizar mapa BCE/Ajustes | DetRetri tiene su propia aritmética (V0.4) y sus propias hojas (93/94) |
| D5 | Validación protegida **por período**: Q1 exige fórmulas HU-07 (F46, F176…, D67…); Q2 exige el mapa T0 (F53…, D73…, F519/F521 excluidos de protegidas-fórmula según T0) | Mantener invariante único | El invariante actual es empíricamente falso en Q2 (V0.5: `F46 debería seguir siendo fórmula`). Parametrizar lo vuelve verdadero sin tocar Q1 |
| D6 | Misma pasada de escritura HU-07..HU-11 (overload lista; un `File.Copy` + validación pre/post con mapa Q2) | Segundo pase / writer separado para DetRetri | La superficie nueva son celdas del mismo modo leaf. Un segundo pase rompería atomicidad y hash A4 |
| D7 | Matcheo estricto `Single` por `Ase.Id` en todo gate nuevo (validador + coherencia DetRetri) | Reutilizar agregados o fallback al primer consolidado | Con 5 consolidados el fallback es bug silencioso (doctrina HU-07 G4) |
| D8 | Golden Q2 = canónico reutilizado (sin re-fijar) + Capa A completa incl. A6 | Nuevo oráculo o "elegir en el test" | Dos oráculos = merge verde en falso; A6 cierra el recorte HU-11 |

### 2.3 Escritura 2.6 ampliada (misma sesión atómica)

```text
File.Copy canónico → salida (una vez, igual que HU-07..HU-11)
  └─► ValidarFormulasProtegidas (mapa Q1 intacto + mapa Q2 T0: visibles R1/R2/R4-Q2
      por período; DetRetri-D escribibles V0.4; DetValiRetri/VALIDACION_* /
      INTERVENTORIA / ANT EXT-REV siempre protegidas;
      ProtegidasBceParaPeriodo(true) verificado contra el canónico — M1)
        └─► EscribirCeldasLeaf HU-07..HU-11 (intactas; solo activas según período)
        └─► EscribirCeldasLeafQ2 (R1/R2/R4 por ASE según mapa T0 + variante ASE5 V0.3)
        └─► EscribirCeldasDetRetriQ2 (enteros V0.4 por ASE)
              └─► revalidar fórmulas protegidas → guardar
```

Ante cualquier fallo: borrar salida parcial (patrón existente). Plantilla origen jamás mutada (hash A4 extendido al mapa Q2). Sin capacidad de filas: fail-fast honesto (HU-07), nunca insert/delete.

### 2.4 Dominio (Core, sin deps)

- Variante R1-Q2-ASE5: el reader leaf resuelve las 2 filas `Mes|Total` de la fuente ASE5-Q2 a slots V0.3 (F531/F552/L531; EXTEMP según T0); si falta un header/slot esperado → fallo que nombra ASE5+R1.
- R1/R2/R4-Q2 ASE1–4: mapas T0 por (`Ase.Id`, período); prohibidos offsets; slot ausente → fallo que nombra ASE+reporte (nunca 0 silencioso; distinguir "leído 0" de "slot ausente").
- DetRetri de dominio: `DetRetri_D(ase) = DetRetriRounder.Round(consolidado.D104_D108(ase))` (V0.4; usa `DetRetriRounder`, prohibido otro redondeo).
- `WorkbookLeafInputs` / detalle DetRetri: extensión mínima aditiva para DetRetri-Q2 por ASE (nullable/default que preserve HU-11 intacto cuando no aplica, patrón `AjustesSfT = null`).
- `ProcesadorPeriodo`: integra lectura Q2 + DetRetri al flujo por ASE en modo Q2 con fail-fast que nombra ASE **y reporte**; modo Q1: pasos Q2 se omiten (mismos asserts que hoy).
- Coherencia `WorkbookLeafCoherence`: `ValidarContraResultado` Q2 itera DetRetri con matcheo estricto `Single` por Id (misma regla HU-07, extendida).

### 2.5 `IValidador` 2.6 ampliada (sin reabrir HU-04..HU-11)

1. Todo lo HU-07..HU-11 intacto (matcheo estricto por `Ase.Id`, `GranTotal = Σ`, gates leaf/empresa/banco/balance/ajustes).
2. Nuevo, **solo si `Periodo.NumeroQuincena == 2`**: por cada ASE, `DetRetri_D(ase)` de dominio == DetRetri fuente/cache-golden ±0.5 (matcheo `Single` por Id; si falta el insumo → error que nombra el ASE).
3. Q1 sin cambios: ningún gate nuevo activo en Q1 (regresión 105/105 + 24/24 ciega a Q2).
4. El validador NO abre `.xlsx`. La composición DetRetri-D va a Capa A (dominio vs cache golden), no a gates de fórmula.

### 2.6 UI — Visual Design Intent (delta mínimo)

Densidad Balanced, mismos GroupBoxes, sin restyle/colores/iconos. El `txtLog` agrega, por cada ASE en modo Q2, una línea DetRetri (origen D104:D108 / entero redondeado, "esperado post-Excel"). Serilog: mismos eventos HU-07..HU-11 con propiedad `Hoja = "DetRetri2026072"`. Sin nuevos controles.

### 2.7 Golden Capa A Q2 completa (honestidad HU-06..HU-11)

| # | Qué | Contra qué | Tol |
|---|---|---|---|
| A1 | Celdas Q2 escritas en la **salida** (R1/R2/R4 por ASE según T0 + DetRetri-D V0.4) | Mismas celdas **leaf** del golden canónico | ±0.5 |
| A2 | DetRetri de **dominio** por ASE (`ROUND(D104:D108,0)`) | Cache golden DetRetri-D Q2 | ±0.5 |
| A3 | DetRetri/DetValiRetri no-escritas + `VALIDACION_*` + INTERVENTORIA + ANT EXT-REV + Q1-protegidas siguen siendo fórmula / intactas en la salida | Estructura | n/a |
| A4 | SHA256 canónico igual antes/después | — | n/a |
| A5 | **Prohibido** comparar cache de fórmula de la salida vs golden; **prohibido** usar agregados HU-02 o totales Q1 como oráculo | — | prohibido |
| A6 | `TotalAse`/`GranTotal` de dominio (con ajustes + R1/R2/R4-Q2 absorbidos) vs cache golden D104:D109 — **cierra el pendiente HU-11** | Cache golden canónico | ±0.5 |
| A7 | Q1 intacto: 78 Q1 + golden Q1 re-assert sin duplicar suite (regresión total 105/105 + 24/24) | Golden Q1 | ±0.5 |
| A8 (M1) | Path Q2 del writer ejercitado: `esQuincena2 = true` + `ProtegidasBceParaPeriodo(true)` matchea hojas/celdas reales del canónico (incl. sheets 93/94) | Estructura canónica | n/a |

Capa B (manual Excel: abrir, recalcular, comparar DetRetri-D + D104:D109 vs canónico) fuera de CI, protocolo §5.3. **Capa B manual queda como acción del usuario.**

### 2.8 File changes previstos

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapQ2.cs` | Crear | Mapa explícito por (`Ase.Id`, período): visibles V0.1 + editables T0 + DetRetri-D V0.4 |
| `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` | Modificar | Dispatch Q2 por período/ASE (incluida variante 2-filas ASE5-Q2 V0.3); path Q1 intacto |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapPorAse.cs` | Modificar | Validación protegida **por período** (Q1 fórmulas HU-07 / Q2 mapa T0, D5) — solo si el diseño lo exige; si no, el dispatch vive en el mapa Q2 |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Escritura leaf-Q2 + DetRetri-Q2 en la misma pasada + mapa protegido por período + M1 |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs` | Modificar | Coherencia DetRetri-vs-consolidado Q2, matcheo estricto |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Modificar | Gates §2.5 (Q2 exige DetRetri; Q1 intacto) |
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` | Modificar | Integrar lectura Q2 + DetRetri al flujo por ASE (fail-fast ASE+reporte) |
| `Remuneracion.Core/Models/WorkbookLeafInputs.cs` (+ detalle DetRetri si T0 lo exige) | Modificar/Crear | Extensión aditiva nullable (patrón `AjustesSfT = null`) |
| `Remuneracion.Core/Interfaces/IWorkbookLeafInputReader.cs` | Modificar | Overloads/variante Q2 si el dispatch lo requiere (contrato aditivo) |
| `Remuneracion.WinForms/Form1.cs` | Modificar | Resumen DetRetri por ASE en log (§2.6) |
| `Remuneracion.IntegrationTests/GoldenDetRetriQ2Tests.cs` | Crear | Capa A completa Q2 (§2.7, canónico HU-11) |
| `Remuneracion.IntegrationTests/DetRetriQ2Tests.cs` | Crear | Dominio: variante 2-filas V0.3, composición V0.4 exacta, mismatch nombra ASE+reporte |
| `Remuneracion.IntegrationTests/ProcesadorPeriodoTests.cs` | Modificar | Procesador Q2 5/5 certificable (levanta el recorte T0-0.6 HU-11) + M1 |
| `Remuneracion.IntegrationTests/Insumos.cs` | Modificar | Helpers `R1Q2Ase5()`, `DetRetriQ2(aseId)`, período Q2 (canónico ya existe) |

**No tocar (salvo bug blocker):** `IPlantillaWriter`/validation-only (HU-04); agregados HU-02; coherencia `F25`-Extemp HU-05; mapas HU-08/HU-09/HU-10; `DetRetriRounder` (se reutiliza); reglas Q1 de `ValidadorBasico`; cadena 2.5 certificada; recorte 1 HU-11 (HU-08 fuera de Q2); `requirements/` legado.

---

## 3. SPEC / Acceptance Criteria (trazabilidad Propuesta §10 + base docs)

### Requirement 1 — Cell-map R1/R2/R4-Q2 por ASE/período con Q1 intacta (CA-1/CA-2; V0.1/V0.2)

El sistema **MUST** leer R1/R2/R4-Q2 de los 5 ASE a los slots del mapa Q2 (T0 + V0.3) y **MUST NOT** alterar el path Q1: la rama Q2 activa solo con `NumeroQuincena == 2`.

- GIVEN fuente R1-Q2-ASE5 → WHEN `LeerLeafInputs(ase5, periodoQ2, …)` → THEN slots V0.3 poblados y TOT_OPT de dominio = 12033011685.71 = D13 golden ±0.5.
- GIVEN cualquier Q1 (78 tests) → THEN comportamiento bit-a-bit idéntico.
- GIVEN slot T0/V0.3 ausente en fuente → THEN fallo que nombra ASE+reporte (fail-fast honesto, sin salida).

### Requirement 2 — DetRetri Q2 en enteros, composición congelada (CA-2/CA-3; V0.4)

El sistema **MUST** calcular `DetRetri_D(ase) = DetRetriRounder.Round(D104:D108(ase))` y escribirlo en la misma pasada. **MUST NOT** crear otra regla de redondeo; **MUST NOT** asumir otro origen (V0.4 probada 5/5, no hipótesis).

### Requirement 3 — Writer multi-ASE Q2 certificable + validación por período + M1 (CA-3/CA-4; V0.5)

El path Q2 del writer **MUST** pasar la validación protegida parametrizada por período (D5) y ejercitarse en CI por primera vez: `esQuincena2 = true` con `ProtegidasBceParaPeriodo(true)` matcheando hojas/celdas reales del canónico (incl. sheets 93/94). Si el `Replace` de sufijo no matchea → bug latente declarado y corregido aquí (mapa parametrizado real, nunca `Replace` ciego).

### Requirement 4 — Procesador Q2 5/5 + Capa A completa incl. A6 (CA-1/CA-3; recorte HU-11)

El procesador Q2 **MUST** producir salida certificada 5/5 (levanta el recorte T0-0.6 HU-11: la salida existe y cuadra) y la matriz §2.7 **MUST** incluir A6 (TotalAse/GranTotal Q2 vs cache canónico ±0.5). Si A6 no cierra → veredicto fallido con recorte declarado (NEEDS_CONTEXT), nunca invención.

### Requirement 5 — Guardas fuera de alcance (CA-4)

`DetValiRetri`/`VALIDACION_*` Q2, `INTERVENTORIA` L25:N31, `ANT EXT-REV` y resto 2.7 **MUST** estar en el mapa protegido (fallan la escritura si alguna deja de ser fórmula). Si la plantilla no alcanza → fail-fast honesto, **MUST NOT** insert/delete de filas.

| CA §10 | HU-12 v2 |
|---|---|
| CA-1 | Lee R1/R2/R4-Q2 de los 5 ASE (fail-fast nombra ASE+reporte; locator agnóstico intacto) |
| CA-2 | Cell-map Q2 + DetRetri = dominio ±0.5 por ASE (`DetRetriRounder`; agregados HU-02 ≠ DetRetri) |
| CA-3 | Capa A Q2 completa incl. A6; DetRetri correcta post-Excel (Capa B residual) |
| CA-4 | Reassert mapa por período + M1 + canónico no mutado (hash) |
| CA-5 | Gates Q2 por quincena + coherencia DetRetri |
| CA-6 | Serilog + log DetRetri por ASE |
| CA-7 | Mismo flujo 5-ASE + resumen DetRetri (Q1 sin cambios visibles) |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 800–1300 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 T0 réplica-HU-07 sobre Q2 + mapa Q2 congelado → PR2 reader Q2 (dispatch + variante ASE5) → PR3 writer Q2 + validación por período + M1 → PR4 validador + período + UI → PR5 golden completa + A6 + regresión |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes. Chained PRs recommended: Yes. Chain strategy: pending.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 0 | Discovery T0 réplica-HU-07 sobre Q2 + mapa Q2 congelado | PR 1 | Bloquea todo; solo lectura + datos. V0.3/V0.4 se reutilizan |
| 1 | Reader Q2 (dispatch período/ASE + variante ASE5) | PR 2 | Depende de PR 1 |
| 2 | Writer Q2 + validación por período + M1 | PR 3 | Depende de PR 1 |
| 3 | Validador Q2 + período + UI | PR 4 | Depende de PR 1 |
| 4 | Golden completa + A6 + regresión 105/105 + 24/24 | PR 5 | Depende de PR 2–4 |

### Phase 0 — Discovery T0 (BLOQUEANTE, solo lectura, réplica del método HU-07/T0 §0.1–0.5 sobre Q2)

- [ ] 0.1 Reutilizar el golden canónico (`Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx`); registrar SHA256 de trabajo. No se re-fija. **V0.1/V0.3/V0.4 se toman como dados** (no se re-dumpean salvo verificación puntual).
- [ ] 0.2 Dumpear por OpenXML (mismo método HU-07: unzip + `<c>` raw) **todas** las celdas operando de cada visible Q2 por ASE y clasificarlas: valor editable vs fórmula. R1: visibles F53/F206/F343/F468/F558 + EXTEMP F55/F208/…/F560 y sus operandos (F552/F531/L531 ya cerrados para ASE5; resto por ASE). R2: visibles `E43…` + operandos por ASE. R4: visibles `D73…` + operandos por ASE. Registrar `<f>` (si existe) y `<v>` de cada una.
- [ ] 0.3 Confirmar EXTEMP Q2 por ASE (candidatos ASE5 F538/F512/L512) + las 2 filas Mes/Total de la fuente ASE5-Q2 (filas 33/54) contra slots V0.3. **Nada de la variante entra al código sin esta confirmación.**
- [ ] 0.4 Dumpear hojas DetRetri/DetValiRetri Q2 del canónico (sheets 93/94): columna D por ASE celda-por-celda (valor-vs-fórmula), refs `-BCE!F3..F7` y `VALIDACION_*`; check de capacidad de filas (si no alcanza: fail-fast honesto, sin insert/delete).
- [ ] 0.5 Verificar M1 a nivel estructura: ¿`ProtegidasBceParaPeriodo(true)` matchea hojas/celdas reales del canónico (incl. 93/94)? Registrar match celda-por-celda o declarar el bug latente del `Replace`.
- [ ] 0.6 Mapeo fuente → template Q2 por bloque (espejo de HU-07 T0-0.3): labels Mes/Total, Subsidio/Contribución, Especiales por ASE en fuentes `REMUNERACION 2026072/{1..5}-*/`; confirmar ASE4-sin-Especiales si aplica en Q2.
- [ ] 0.7 Congelar `WorkbookLeafCellMapQ2` (visibles V0.1 + editables T0-0.2 + DetRetri-D V0.4 + protegidas incl. DetValiRetri/VALIDACION_*/INTERVENTORIA/ANT EXT-REV/F46-Q1-excluidos-en-Q2). **Nada entra al código sin esta tabla.**
- [ ] 0.8 Confirmar baseline Q1 intacto (78 Q1 + D DetRetri-Q1 + mapa HU-07 sin cambios).

### Phase 1 — Dominio (DetRetri + dispatch)

- [ ] 1.1 Extensión aditiva para DetRetri-Q2 por ASE (nullable/default HU-11 intacto) + propiedad calculada `Round(D104:D108)` (V0.4; usa `DetRetriRounder`; prohibido otro redondeo).
- [ ] 1.2 Tests in-memory: `ROUND(D104:D108,0)` cuadra vs golden ±0.5 por ASE (5/5); Q1 idéntico; origen-equivocado (p. ej. agregado HU-02 como DetRetri) falla a propósito.

### Phase 2 — Lectura (dispatch Q2 + variante ASE5)

- [ ] 2.1 Dispatch por (`Ase.Id`, `NumeroQuincena == 2`) en el reader leaf R1/R2/R4 (mapa Q2 T0-0.7; prohibidos offsets); path Q1 intacto.
- [ ] 2.2 Variante ASE5-Q2 con slots V0.3 confirmados en T0-0.3; fail-fast que nombra ASE5+R1 si falta un slot (nunca 0 silencioso; distinguir "leído 0" de "slot ausente").
- [ ] 2.3 `Insumos.cs`: helpers `R1Q2Ase5()`, `DetRetriQ2(aseId)` (canónico ya existe).

### Phase 3 — Orquestación + UI delta mínimo

- [ ] 3.1 `OpenXmlPlantillaWriter`: escritura leaf-Q2 + DetRetri-Q2 (V0.4) en la misma pasada + validación protegida por período (D5) + M1.
- [ ] 3.2 `ValidadorBasico` + `WorkbookLeafCoherence`: gates §2.5 (Q2 exige DetRetri + matcheo estricto; Q1 intacto).
- [ ] 3.3 `ProcesadorPeriodo`: integrar lectura Q2 + DetRetri al flujo Q2 por ASE (fail-fast ASE+reporte); Q1 omite Q2.
- [ ] 3.4 `Form1`: resumen DetRetri por ASE en log (§2.6) + Serilog `Hoja = "DetRetri2026072"`.

### Phase 4 — Pruebas y evidencia

- [ ] 4.1 `DetRetriQ2Tests` (in-memory + fuentes Q2 reales sin salida): variante 2-filas V0.3 ok, composición V0.4 exacta, mismatch nombra ASE+reporte, Q1 intacto.
- [ ] 4.2 `GoldenDetRetriQ2Tests`: matriz §2.7 completa (5 ASE, canónico; A6 incluido — cierra el pendiente HU-11).
- [ ] 4.3 **OBLIGATORIO M1**: writer Q2 ejercitado (`esQuincena2 = true`, sufijo real verificado; el `Replace` ciego queda prohibido por test).
- [ ] 4.4 `ProcesadorPeriodoTests`: procesador Q2 5/5 certificable (levanta el recorte T0-0.6 HU-11; la salida existe y cuadra).
- [ ] 4.5 Negativas: slot T0 ausente en ASE5 (falla ASE5+R1); plantilla sin filas (fail-fast, sin insert/delete); salida == plantilla (no in-place); overwrite cancelado; Q2 con rama Q1 (bloqueado, intacto).
- [ ] 4.6 Los 105 tests + harness 24/24 existentes verdes; build 0 warnings; CRLF; sin commit.

### Phase 5 — Documental

- [ ] 5.1 Capa B manual §5.3 ejecutada una vez sobre el canónico y evidenciada (acción del usuario; sin fingirla como gate de merge).
- [ ] 5.2 Cierre deja explícito el frente 2.7 (siguiente HU propuesta: validaciones cruzadas).

---

## 5. Test Strategy

| Capa | Qué | Cómo |
|---|---|---|
| Unidad | Dispatch Q2 (slots T0/V0.3, fail-fast ASE+reporte) | Fuentes Q2 reales, sin Excel de salida |
| Unidad | Composición DetRetri V0.4 (`DetRetriRounder`, Q1 intacto, `Single` por Id) | In-memory |
| Unidad | Validación por período (Q1 fórmulas HU-07 / Q2 mapa T0) | In-memory + estructura |
| Integración | Período Q2 5/5 (salida temp, fail-fast ASE+reporte) | Insumos Q2, canónico |
| Golden Capa A Q2 | Matriz §2.7 completa incl. A6 + A8(M1) | OpenXML read-only + aritmética dominio (A5 aplica) |
| Regresión | 105/105 + harness 24/24 Q1 verdes | Suite existente, sin cambios |
| UI | Resumen DetRetri | Funcional manual (sin harness) |
| Capa B | DetRetri-D + D104:D109 post-Excel | Manual — §5.3 |

### 5.1 Fixtures

- Golden/plantilla Q2: canónico (sin re-fijar).
- Fuentes Q2: `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/` (R1 ASE5 rango 16072026–31072026 con 2 filas Mes/Total filas 33/54; resto para paridad).
- Regresión: golden + fuentes Q1 (intactos).
- Referencia: V0.3 (12033011685.71 = D13) + V0.4 (ROUND D104:D108 5/5) + tablas T0-0.7, tolerancia ±0.5.

### 5.2 Casos negativos obligatorios (nombran ASE y reporte)

Slot T0/V0.3 ausente en R1-ASE5 (falla ASE5+R1); `Replace` de sufijo que no matchea (M1 lo demuestra fallando); origen-equivocado como DetRetri (prohibido — el test lo demuestra fallando); salida == plantilla (no in-place); plantilla sin filas DetRetri (fail-fast honesto, sin insert/delete); Q1 con rama Q2 activa (bloqueado, intacto); HU-08 en Q2 (fuera de alcance — el test no lo toca).

### 5.3 Protocolo manual Capa B (no CI — acción del usuario)

1. Generar salida Q2 a ruta distinta del canónico. 2. Abrir en Excel, recalcular. 3. Comparar DetRetri-D por ASE + D104:D108 + D109 vs canónico ±0.5. 4. Evidencia adjunta. No es gate de merge.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | Q2 = extensión de lectura/escritura (mapa hermano + ramas); DetRetri = extensión de datos (mapa + gates); cálculo/período/writer conservan su rol HU-07..HU-11. |
| **O** | Se agregan mapa/rama/gates por período; el path Q1 y la cadena 2.5 funcionan sin modificarse (abierto sin modificar). |
| **L** | El reader leaf suma rama Q2 sin cambiar el comportamiento Q1 ni el fail-fast existente en el resto. |
| **I** | Detalle DetRetri-Q2 separado de `BalanceScInputs`/`AjustesSfTInputs`; reader/validador/writer crecen por overload/rama. |
| **D** | Core define mapas/gates por período; Infrastructure/WinForms componen. Sin nuevas deps; `DetRetriRounder` reutilizado. |

### 6.2 Best Practices

- La verdad del workbook manda: Q1 y Q2 tienen layouts distintos (V0.1/V0.2); mapas hermanos por período, no unificación forzada (Rector §6 preservación).
- V0.3/V0.4 probadas contra el golden se reutilizan; lo no cubierto por V0 lo dicta T0 (honestidad §0.2).
- Mapa explícito por `Ase.Id` + período verificado, no offsets ni filas/columnas fijas (Rector §11.3).
- `DetRetriRounder` único y existente (V0.7); prohibido duplicar la regla de redondeo.
- Una escritura atómica; canónico nunca mutado; hash A4 extendido.
- Golden honesto Q2 con A6 cerrado (OpenXML no recalcula; A5).
- Fail-fast nombra ASE+reporte; sin salida certificada ante fallo; sin insert/delete de filas.
- Q1 blindada por construcción (dispatch por período + 105/105 + 24/24 como red) y M1 deja de ser deuda (path Q2 ejercitado).

### 6.3 Performance

- 5 ASE × lecturas Q2 + ~5 escrituras leaf-Q2 + 5 DetRetri en la misma sesión OpenXML. Irrelevante a esta escala; `Task.Run` existente para no congelar el form.

**Veredicto:** APROBADO como it. 2.6 ampliada de Fase 2 **si** T0 congela el mapa Q2 con evidencia réplica-HU-07 (§4 Fase 0: V0.2 pendiente) y se acepta CA-3 parcial (Capa A en CI, Capa B manual como acción del usuario).

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| R1/R2/R4-Q2 ASE1–4 traen layouts que el mapa T0 no puede absorber sin insert/delete | Media | Fail-fast honesto (HU-07) + NEEDS_CONTEXT con recorte, nunca insert/delete |
| EXTEMP Q2 por ASE (incl. F538/F512/L512 ASE5) no cierra contra el golden | Media | T0-0.3 lo confirma o declara veredicto fallido parcial; prohibido inventar slot |
| R2/R4-Q2 refs por ASE divergen del patrón `E43…`/`D73…` esperado | Baja | T0-0.2 dumpea la verdad; el mapa sigue al dump, no a la expectativa |
| `Replace` de sufijo M1 esconde hojas inexistentes en Q2 | Media | T0-0.5 + test A8 obligatorio; mapa parametrizado real si el `Replace` no matchea |
| Ceros legítimos Q2 confundidos con "falta de lectura" | Alta | El reader distingue "leído 0" de "slot ausente" (fallo); tests de ceros explícitos |
| Inflar a 2.7 dentro de esta HU (DetValiRetri/VALIDACION_* tientan) | Media | §0.3 out-of-scope + protegidas, no implementadas; rechazar PRs que lo metan |
| Comparar cache de salida vs golden y "cerrar" A6 en falso | Alta | A5 + canónico único; este plan lo prohíbe |
| Q1 regresión rota por el dispatch de período | Media | 105/105 + 24/24 como red en PR5; Q1 intacto por construcción (G5/D1/D5) |
| Reabrir recorte 1 (HU-08 en Q2) o cadena 2.5 por inercia | Media | Fuera de alcance explícito (V0.6); rechazar PRs que lo metan |

---

## 8. Rollback

- Revertir PRs en orden inverso (golden → período/UI/validador → writer → reader Q2 → mapa Q2/T0).
- HU-04..HU-11 intactas sin esta HU: rama Q2 inactiva en Q1; Q2 sigue con recorte T0-0.6 honesto (sin salida certificada 5/5).
- `Docs/Insumos/` untracked: jamás como destino; pisada se restaura desde backup.
- Si T0 demuestra layout que exige insert/delete o EXTEMP que no cierra (Riesgos 1–2), recorte con rebase a lo verificado (V0.3/V0.4 + ASE cubiertos), no invención.

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §9 it. 2.6 / §6 / §10 (±0.5) + Detalle de plantilla y Prompt Maestro como origen funcional (§1.5). 2.7 lógica, INTERVENTORIA, ANT EXT-REV quedan fuera. Proceso de Recaudo NO aplica salvo cita directa.
2. Cell-map Q2 hermano del mapa HU-07, explícito por (`Ase.Id`, período) con dispatch por período (D1); path Q1 intacto; prohibidos offsets.
3. Variante R1-Q2-ASE5 con slots V0.3 (F531/F552/L531; EXTEMP a confirmar T0-0.3); fail-fast ASE5+R1 ante slot ausente.
4. DetRetri-Q2 con composición congelada V0.4 (`ROUND(D104:D108,0)` vía `DetRetriRounder`); `DetValiRetri`/`VALIDACION_*` = protegidas.
5. Validación protegida parametrizada por período (D5: Q1 = fórmulas HU-07; Q2 = mapa T0); M1 verificado contra el canónico incl. sheets 36/37 (fe de erratas HU-17 V5: el borrador decía "sheets 93/94"; la verdad canónica, confirmada por revisión HU-12 con `workbook.xml` y re-verificada en HU-17 con `mcp-excel` sobre ambos goldens, es que `DetRetri202607{1,2}`/`DetValiRetri202607{1,2}` son las hojas 36/37 de 40).
6. Una sola escritura atómica; sin insert/delete de filas (fail-fast honesto si la plantilla no alcanza).
7. Golden Q2 = canónico reutilizado (no se re-fija); Capa A completa incl. A6 cierra el recorte HU-11; Capa B manual residual como acción del usuario.
8. Recorte 1 HU-11 (HU-08 fuera de Q2) y cadena 2.5 certificada se mantienen intactos (V0.6).
9. UI delta mínimo + Serilog DetRetri por ASE; OPA = Ejecutar.
10. Apply espera aprobación + PRs encadenados con T0 al frente (Unidad 0–4, §4).
11. Si algo no cierra ±0.5: NEEDS_CONTEXT con recorte, nunca invención.
12. **M2 (nota de aceptación, HU-17 V6):** `TieneColumnaEspeciales = false` fijado por código para las fuentes de saldos-nota/retribución (evidencia T0-0.5: 5/5 sin columna "Especiales"); ausente = 0. No es un fallo: es el comportamiento esperado, documentado en el manual y el instructivo Capa B.

---

## 10. Lecciones del plan v1 (premisa refutada — se conserva, no se borra)

La v1 ("DetRetri Q2 + absorción layout R1-Q2-ASE5") asumía implícitamente que el mapa HU-07/Q1 seguía válido en Q2 salvo el hueco ASE5 (recorte T0-0.6: 2 filas vs 3) y un veredicto F519/F521. **Refutada por V0.1/V0.2/V0.5:**

1. El mapa Q1 es inválido en Q2 en su totalidad: visibles D9:D13 → F53/F206/F343/F468/F558 (no F46/…), D66 → R4!D73 (no D67), y F46/F176/…/F521 son valores 0 o inexistentes en el template Q2.
2. El writer Q2 no muere por el recorte ASE5 sino **antes**, en la validación protegida R1 (`F46 debería seguir siendo fórmula`) — invariante Q1 falso en Q2.
3. La composición DetRetri-Q2 no requería descubrimiento (V0.4 la cierra 5/5: `ROUND(D104:D108,0)`); lo que faltaba era el cell-map Q2 completo que la alimenta.
4. El doble desenlace F519/F521 de la v1 (D3) queda **derogado**: en Q2 esas celdas no son el visible; el mapa Q2 las trata según T0-0.7, no como veredicto estructural.
5. Lección de método: ante un cambio de quincena con template distinto, el primer T0 debe dumpear los **visibles CONSOLIDADO** (sheet18) antes de asumir continuidad de layout — el orquestador lo hizo y reorientó el alcance a la opción (a).

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
