# Plan 52 — Mapeo por sección de los operandos empresa R1 (fin del índice de ocurrencia)

| Campo | Valor |
|---|---|
| Estado | **CERRADO CON STOP 2026-10-22** (H0 rojo: la regla por sección es correcta, pero la heurística ordinal no basta; rediseño en Plan 53) |
| Supersesión | **2026-10-22:** la Ruta B (§13) se implementa en `plans/53 - contrato-seccion-por-operando-r1.md` (**APROBADO**, H1/H2 cerrados). El camino ordinal que este plan dejó detenido queda **prohibido** y retirado del reader (`ExtraerCeldasR1` resuelve por sección declarada por celda; `OrderByDescending(|F|)` eliminado de ese camino). Este documento se conserva como diagnóstico histórico (T0: `Docs/T0-Plan52-Seccion-Operandos-Empresa-R1.md`). |
| Fecha | 2026-10-22 |
| Base de regresión | `p51-h2-suite.trx` (**1260 / 1193 / 67 / 0**) + `p51-h2-setdiff.txt` |
| Familia | Opcionalidad de fuentes/canales (residuo `R-PARCIAL-OCT-ASE3-R1-OCC` del Plan 50, reclasificado 2026-10-22) |
| Toca | `ExtraerCeldasR1` (reader) + doc-comments del enum `FuenteR1` (Core, solo comentarios) |
| No toca | Plantillas v1/v2 (SHA §8), writer, mapas R2/R4, banco, Plan 50/51 y sus deudas, `CalcularVisibleR1` (lógica) |

---

## 0. Clarification Gate (pre-respondido por el encargo — se registra, no se re-pregunta)

El encargo del Ingeniero ya arbitra los 5 puntos que normalmente bloquearían el plan; se dejan por escrito para que el §12 solo pida lo que falta:

1. **Alcance cerrado A1–A4** (§3): solo operandos `EditablesR1PorEmpresa`; R2/R4 solo se declaran (D6), no se tocan.
2. **Reutilizar, no inventar** (§5-D1): predicados existentes (`SeccionPorCierre` / `PorFronteraArriba` / `CanalEfectivo`); cero lógica nueva de sección.
3. **T0 read-only bloqueante** (§10-T0): matriz ASE × empresa × período + arbitraje del oráculo manual; si una sección no es tipable con los predicados existentes → **STOP** (no inventar).
4. **Tests no se debilitan**: 0 `Skip`, 0 `Assert.True(true)`, 0 borrados. El throw-test vigente se **reescribe** a aserciones de valor (baja+alta declarada, §9), no se elimina.
5. **Rojo fuera de lista = STOP** (§9): el set-diff de H2 debe contener EXACTAMENTE las altas/bajas cerradas en §9; cualquier otra divergencia detiene el plan y escala.

---

## FASE 1 — PROPOSE

### 1. Contexto y doctrina aplicable (no se re-discute)

- **Reglas 6 y 7 del project-context** (corregidas 2026-10-22, reunión con la propietaria): las fuentes **Y** los canales pueden venir o no venir; una sección/canal ausente **no es estructura rota**; su celda destino queda en **0**. Prohibido resolver por ocurrencia/índice (`filas[count-1]`); la señal es la **frontera** (`Mes`, `Aplicacion`, `Componente`).
- **Plan 50 §8, residuo 1** (`R-PARCIAL-OCT-ASE3-R1-OCC`, reclasificado 2026-10-22): NO es estructura rota — la fuente OCCIDENTE no trae la primera sección (`AFaseo`) en octubre-1 y eso es normal. El defecto es el **mapa por ocurrencia**. Fix en plan propio (= este plan).
- **Plan 51** (misma familia de opcionalidad, otro reporte): precedente de método — resolver por **etiqueta observada** (header-scan), ausente = omitir con huella, sin ramas por período/empresa (D-H). Este plan aplica el mismo método al eje sección.
- **Planes 41/45/46** (`SeccionEspejoR1` / `ContratoSeccionPeriodFreeR1` / canal efectivo): contrato de sección vigente y reutilizable. **Plan 46 supersede al 45** en tipado (cierre + canal efectivo); Plan 41 §E1-1/E1-3 (prohibición de fallback posicional, `ERR-SIN-SLOT`) sigue vigente.

### 2. Problema (evidencia probada con bytes — no re-investigar)

`EditablesR1PorEmpresa[(4,3)]` (OCCIDENTE × ASE3) resuelve por **índice de ocurrencia** sobre listas ordenadas por `|F|` descendente (`ExcelDataReaderWorkbookLeafInputReader.cs:1585-1637`):

| Operando | Resolución actual | Supuesto oculto |
|---|---|---|
| `TotalMain` / `EspecialesMain` | `filasTotal[0]` | el mayor `\|F\|` es el bloque oportuno |
| `TotalMenor` / `SubsMenor` / `EspecialesMenor` | `filasTotal[count-1]` / `filasSubs[count-1]` | el menor `\|F\|` es la primera sección |
| `TotalMedio` | `filasTotal[1]` | siempre hay ≥ 2 filas |

Medición por período (OCCIDENTE antes de `A=Componente`, Plan 50 T0 §2.2 + encargo):

| Período | Filas | Contenido | Efecto del código actual |
|---|---|---|---|
| Julio-2 ASE3 | 1 | `−16.458,46` (primera sección) | `Menor = [0]` duplica; `Medio = [1]` → **throw** |
| Agosto-2 ASE3 | 2 | `−522,80` (primera) + `897.711.004,34` (oportuno) | `[0]`=oportuno, `[1]`=primera: coincide por accidente de magnitud |
| Octubre-1 ASE3 | 1 | `286.711.907,53` (oportuno, R23) | `Medio = [1]` → **fail-fast** (`S2_1_ParcialConLabelPresente`, Plan 50); si se "corrigiera" con `count-1`, `TotalMenor = [0]` duplicaría el bloque grande en `F227 = F253+F237+F216−L237` |

**Causa raíz:** la estructura no es estable (1 / 2 / 1 filas según el período) y el orden por magnitud no es una señal de sección: con 1 fila, `count-1 == 0` colapsa `Menor` sobre `Main` (duplicado silencioso si no hubiera throw), y `Medio` exige una cardinalidad que la fuente no promete. Resolver por ocurrencia es el defecto; la regla 6 lo prohíbe explícitamente.

Tensión aparente que el T0 debe cerrar (§10-S1.5): julio-2 ASE3 figura con 1 fila y julio-2 es línea verde. El T0 mide el conteo exacto con firma A–E completa (el encargo cuenta filas `Total`; el reader cuenta `C=label ∧ D=Total` con cualquier `E`): si julio-2 tiene 1 fila-empresa, la corrida Q2 multi-ASE no lo ve (recorte HU-20/G2-D1: la conciliación por empresa se omite en Q2) y el single-ASE sí debe medirse. Nada se asume: el T0 lo congela.

### 3. Alcance APROBADO (no ampliar)

- **A1.** Resolver los operandos de `EditablesR1PorEmpresa` por **(firma + sección + canal)**, no por índice. Reutilizar predicados existentes — **cero lógica nueva de sección**. `TotalMain`/`SubsMain`/`EspecialesMain` = bloque de la sección oportuna presente de mayor ordinal; `TotalMenor`/`TotalMedio`/`SubsMenor`/`EspecialesMenor` = secciones presente de menor ordinal / intermedia; **ausente → 0**, nunca `count-1`. Si la sección existe pero el canal/concepto no → 0 (misma regla).
- **A2.** Ningún duplicado por ausencia (regla dura): un operando no puede tomar el valor de otro bloque. Test obligatorio con 1 fila: `TotalMenor = 0` y `TotalMain = 286.711.907,53` (no ambos iguales).
- **A3.** Invariantes intactos: `A='Componente' B='Total'`, `A='Subs/Cont' B='Total'`, `A='Total'` siguen fail-fast si faltan (estructura del reporte). La ausencia de fuente/canal/sección **no** es fail-fast.
- **A4.** No tocar plantillas ni writer. v2 SHA `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB`; v1 SHA `3086AC317595C2F8E84ED24B1C5EE3E52565DD5194E3062878E7FA49C5214`. Values-only.

### 4. No-objetivos (STOP si aparecen)

1. No tocar el Plan 50/51 ni sus deudas (`R-SIGMA-Q2`, `R-MAPA-V2-R2-ASE4`, F-NEW-1).
2. No tocar el banco (`ReporteBanco*`), ni los mapas `EditablesR2PorEmpresa` / `EditablesR4PorEmpresa` (D6 solo declara).
3. No renombrar/reordenar el enum `FuenteR1` (contrato: lo consumen mapa, reader-switch, `CalcularVisibleR1` y tests). Solo se reescriben sus **doc-comments** (semántica por sección).
4. No cambiar columnas de lectura (`F`=col 5, `Especiales`=col 11 fijas): no son el defecto; el header-scan de componentes es frente de otro plan.
5. No reintroducir fallback posicional (Plan 41 E1-1) bajo ninguna forma.

---

## FASE 2 — DESIGN

### D1. Qué predicado de sección usar (y por qué)

**Decisión: `ContratoSeccionPeriodFreeR1.SeccionPorCierre` como tipador primario; `CanalEfectivo` donde aplique; `SeccionEspejoR1.PorFronteraArriba` solo como el fallback que ya trae `SeccionPorCierre` para la cola sin frontera.**

Razonamiento:

1. Las filas-empresa (`C=label ∧ D=Total`, `indice < indiceComponente`) son **filas de datos, nunca fronteras**: su firma es disjunta de `EsMesTotal` (`B=Mes ∧ C=Total`) y `EsAplicacionTotal` (`B∼Aplicacion ∧ C=Total`). Para una fila estrictamente interior, cierre y apertura coinciden (la sección entre dos fronteras consecutivas es la misma por ambos lados), así que `SeccionPorCierre` y `PorFronteraArriba` **acuerdan por construcción**; el T0 lo verifica fila a fila (S1.1) y cualquier desacuerdo es STOP.
2. `SeccionPorCierre` (Plan 46, vigente) es el único que **tipa la primera sección por su frontera de cierre** (`Mes→Oportuno#k`, `Aplicacion→Aplicacion#k`) en vez del `Primera` genérico. Esa distinción es la que permite responder D3 sin magnitud: una "primera" que cierra en `Mes` es `Oportuno#1`; una que cierra en `Aplicacion` es `Aplicacion#1`.
3. `CanalEfectivo` devuelve `null` para estas filas (traen `D=Total` propio, no vacío) — y está bien: aquí la empresa ya viene discriminada por la **firma** (`C=label`), así que la llave efectiva es **(firma-empresa + sección-por-cierre)**. No se crea variante nueva del predicado; se documenta la degradación.
4. `Alinear` (monotónico fuente→oficial) **no se usa** en este camino: alinea secciones del espejo interior, no bloques-empresa del lector de conciliación. Citarlo como referencia, no invocarlo.
5. El reader trabaja hoy sobre `List<object?[]>` crudas y `SeccionPorCierre` exige `IReadOnlyList<FilaEspejoR1>`: el SPEC (§7-H1.2) reutiliza la **construcción existente** de `FilaEspejoR1` (la que alimenta `LeerEspejoR1`), prohibido reimplementar el scan de fronteras.

### D2. Cómo se nombra cada operando (sin romper el contrato del enum)

**Decisión: nombres intactos; se reescribe la semántica en los doc-comments (único cambio en Core).** Tabla normativa (T0 la confirma o STOP):

| Miembro | Semántica nueva (por sección, sin magnitud) |
|---|---|
| `TotalMain` | `F` de la fila-empresa en la sección **oportuna presente de mayor ordinal**; 0 si la empresa no trae bloque oportuno |
| `EspecialesMain` | `Especiales` de esa misma fila; 0 ídem |
| `SubsMain` | `F` de la fila-subsidio de la empresa en la sección oportuna de mayor ordinal; 0 si ausente |
| `TotalMenor` | `F` de la fila-empresa en la sección presente de **menor ordinal**; 0 si solo hay una sección (nunca `count-1`) |
| `EspecialesMenor` / `SubsMenor` | análogos sobre sus filas; 0 si ausente |
| `TotalMedio` / `EspecialesMedio` | `F`/`Especiales` de la fila-empresa en la sección **intermedia** (solo si hay ≥ 3 secciones presentes); **0 si hay < 3** |
| `Block5Total` | intacto (canal `D=5`, sin sección) |

Consecuencia deliberada (va a la lista cerrada §9): con 2 secciones, `TotalMedio` pasa de "duplicar el menor" (índice 1 == `count-1`) a **0**. Los Δ exactos por celda los congela el T0 (S1.7) y los arbitra el oráculo manual.

### D3. Qué pasa con `TotalMain` cuando la primera sección es la única presente

**Decisión: se define por sección, no por magnitud — única fila en primera sección ⇒ `Menor` (y `Main = 0`, `Medio = 0`).**

- Julio-2 ASE3 (única fila `−16.458,46`, primera sección): `TotalMenor = −16.458,46`, `TotalMain = 0`, `TotalMedio = 0` (+ `Subs*` análogos en 0).
- Octubre-1 ASE3 (única fila `286.711.907,53`, oportuno R23): `TotalMain = 286.711.907,53`, `TotalMenor = 0`, `TotalMedio = 0`.
- Regla general: `Main`/`Menor`/`Medio` = ordinal **mayor / menor / intermedio** de las secciones presentes para esa empresa (Totales y Subs por separado); empate de sección (2 filas en la misma sección) o desalineación magnitud≠ordinal en un caso verde actual → **STOP + escala** (no hay regla de desempate aprobada).

### D4. Impacto en los 5 ASE y lista cerrada de altas/bajas

**Decisión: identidad bit-exacta en todo lo verde actual salvo la lista cerrada que congela el T0.** Invariante de no-regresión: para cada (ASE × empresa) donde el código actual **devuelve valores sin throw** en jul-1 / jul-2 / ago-2 / oct-1, la resolución por sección debe reproducirlos **exactos** (igualdad `decimal`, no ±0.5: es el mismo origen de datos). El T0 lo demuestra por construcción (S1.5); si un solo caso verde diverge → STOP.

Altas/bajas esperadas (cerradas, §9): solo (i) el caso oct-1 ASE3 que hoy hace throw (pasa a valores por sección + 0s), (ii) los `TotalMedio` de 2-secciones que hoy duplican al menor (pasan a 0, con Δ congelado por el T0 y arbitrado por el manual), (iii) el posible caso jul-2 ASE3 si el T0 confirma 1 fila (cambio Menor/Main según D3). Todo lo demás: set-diff vacío.

### D5. Test de no-duplicación (obligatorio)

**Decisión: dos tests con insumos REALES por ruta explícita (prohibido `Docs/Insumos` raíz), sin fixtures sintéticos para el caso principal:**

1. `NoDuplicado_Oct1_Ase3_Occidente_UnaFilaOportuna`: `TotalMain == 286.711.907,53`, `TotalMenor == 0`, `TotalMedio == 0`, `SubsMenor == 0`, `EspecialesMenor == 0`, y aserción explícita `TotalMenor != TotalMain` (la que hoy fallaría con `count-1`). Valores `SubsMain`/`EspecialesMain` congelados por el T0.
2. `NoDuplicado_Jul2_Ase3_Occidente_UnaFilaPrimera` (si el T0 confirma 1 fila): `TotalMenor == −16.458,46`, `TotalMain == 0`, `TotalMedio == 0`, `TotalMenor != TotalMain`.
3. El throw-test vigente `S2_1_ParcialConLabelPresente_SigueSiendoFailFast` (`FuentesAusentesPlan50Tests.cs:122-132`) se **reescribe** al nuevo comportamiento especificado (no se borra el método: cambia el cuerpo a aserciones de valor + conserva el comentario de trazabilidad a este plan). Cuenta como 1 baja + 1 alta declaradas.

### D6. Alcance R2/R4 (declarar, no ampliar)

**Decisión: el T0 inventaría R2/R4 (solo conteos, read-only) y el plan declara sin tocar:**

- **R2 tiene el mismo defecto en potencia**: `Total = [0]` (mayor `|E|`) y `SubsCont = [count-1]` colapsan con 1 sola fila-empresa (`Total == SubsCont` ⇒ visible `Total+SubsCont−Especiales` duplica). Si el inventario encuentra un caso real en los 4 períodos → **residuo declarado con dueño** (sigue este mismo patrón en plan hijo con su T0; NO se arregla aquí).
- **R4 no tiene el defecto**: una sola fila `A=label ∧ B=Total` por `FirstOrDefault`, sin índice de ocurrencia (`ExtraerCeldasR4`, `*.cs:1718-1734`). Nada que declarar salvo que el inventario lo desmienta (entonces STOP).
- El control de etiquetas (Plan 50 A3) y `FuentesAusentes` siguen intactos y sin cambios.

---

## FASE 3 — SPEC

### 6. Contratos conductuales (viejo → nuevo)

| Situación | Hoy | Plan 52 |
|---|---|---|
| Empresa con bloques en secciones S₁<…<Sₖ (Totales y Subs por separado) | `Main=[0]`, `Menor=[count-1]`, `Medio=[1]` por `\|F\|` | `Main`=Sₖ oportuna, `Menor`=S₁, `Medio`=intermedia (≥3) |
| Sección requerida ausente (`Menor`/`Medio`/`SubsMenor`/`Esp*` sin fila) | throw (`ValorDe`/`EspecialesDe`) o duplicado (`count-1==0`) | **0 explícito**, sin throw, sin huella nueva (la huella `FuentesAusentes` solo cubre ausencia total del label — intacta) |
| 1 sola fila oportuna (oct-1 ASE3) | throw en `TotalMedio` | `Main`=valor, resto 0 (A2 verificable) |
| 1 sola fila primera (jul-2 ASE3, si T0 lo confirma) | throw/duplicado | `Menor`=valor, `Main=Medio=0` |
| 2 secciones (ago-2 ASE3 y análogos) | `Medio == Menor` (mismo índice 1) | `Medio = 0`; `Main`/`Menor` idénticos a hoy (S1.5) |
| Invariantes `Componente` / `Subs/Cont` / `Total` ausentes | fail-fast | **fail-fast intacto** (mensajes y tipo intactos) |
| Ausencia total del label | 0 + huella (Plan 50) | intacto, sin cambios |
| `Block5Total`, columnas 5/11, `CalcularVisibleR1`, writer, plantillas | — | intactos |

### 7. Cambios archivo por archivo

- **H1.1 (Core, nuevo puro): `ResolvedorOperandosEmpresaR1`** (`Remuneracion.Core/Models/`, BCL puro, D-H: sin I/O, sin período, sin empresa): `IReadOnlyList<(FilaEspejoR1 Fila, int Indice, SeccionFronteraArriba Seccion)> × operando → (fila | ausente)`. Implementa EXACTAMENTE la tabla D2 sobre secciones ya tipadas (ord mayor/menor/intermedia por lista de Totales y de Subs). Sin predicados nuevos: recibe secciones resueltas. + unit tests sintéticos de la regla (ordinales, ausencias, empate→STOP declarado como `null`+código que el reader convierte en throw accionable nombrando ASE+empresa+sección).
- **H1.2 (Infrastructure): `ExtraerCeldasR1`** (`ExcelDataReaderWorkbookLeafInputReader.cs:1560-1640`): conserva firma, invariante `indiceComponente`, `EsFilaEmpresa`, `Bloque5Total`, rama Plan-50 de ausencia total y `ValorDe/EspecialesDe` (ahora solo se llaman con fila existente — el throw por índice desaparece por construcción, el método queda como defensa). Cambia SOLO la resolución: guarda `(fila, indice)` en vez de solo `fila`; construye/reutiliza la lista `FilaEspejoR1` del sheet (la misma que alimenta `LeerEspejoR1` — extraer conversor compartido si no existe punto único); tipa cada fila-empresa con `SeccionPorCierre`; delega la elección al resolvedor Core; ausente → `0m`. `OrderByDescending(|F|)` **se elimina** de este camino (es el defecto).
- **H1.3 (Core, solo comentarios): enum `FuenteR1`** (`WorkbookLeafCellMapPorEmpresa.cs:24-56`): reescribir los 9 doc-comments con la semántica D2 (citar `SeccionPorCierre` + este plan). Cero cambios de miembros, orden o mapa de celdas (`:83-128` intacto).
- **H1.4 (Tests):** fixture T0 read-only `MapeoSeccionOperandosEmpresaR1T0Tests` + doc `Docs/T0-Plan52-Seccion-Operandos-Empresa-R1.md`; reescritura del throw-test S2_1 (baja+alta declarada); tests D5 (no-duplicación, 2 REALES); unit tests del resolvedor; suite + set-diff vs `p51-h2-suite.trx`.

### 8. Invariantes y gates de cierre

- Build 0 warnings; CRLF; sin `git commit/push`; SHAs v1/v2 bit-idénticos antes/después (verificar por hash en H2).
- `grep` de cierre: ningún `ElementAtOrDefault`/`[count-1]`/`OrderByDescending` restante en el camino `ExtraerCeldasR1`; ningún `if (empresaId`/`aseId` nuevo (D-H); ningún `<f>` tocado (el reader no escribe fórmulas por construcción).
- Criterios de aceptación: (i) T0 verde con matriz completa + arbitraje del oráculo; (ii) H1 implementa D1–D6 sin STOP pendiente; (iii) H2 set-diff == lista §9 exacta (fuera de lista = STOP, sin excepciones); (iv) code-review **APROBADO** (0 bloqueantes); (v) SHAs + `git diff --check` limpios.

---

## FASE 4 — TASKS

### T0 (H0, read-only bloqueante — sin T0 verde no se abre H1)

Fixture: `Remuneracion.IntegrationTests/MapeoSeccionOperandosEmpresaR1T0Tests.cs` (solo lectura por rutas explícitas `Docs/Prueba {Julio-1, Julio-2, Agosto-2, Octubre-1}/Insumos/{1-Promoambiental,…}`, nunca `Insumos.Raiz()`). Evidencia: `Docs/T0-Plan52-Seccion-Operandos-Empresa-R1.md`.

| ID | Tarea | Hecho en |
|---|---|---|
| S1.1 | Matriz ASE × empresa × período (jul-1, jul-2, ago-2, oct-1) de filas-empresa R1: conteo `filasTotal`/`filasSubs`, firma A–E completa, valor F, sección por `SeccionPorCierre` Y por `PorFronteraArriba`, frontera que la abre/cierra. Incluye ASE4/ASE5 y casos 1/2/3+ filas | T0 §1 + fixture |
| S1.2 | Confirmar qué sección alimenta cada operando donde hay 2+ filas (ago-2 y análogos): ordinal vs magnitud, fila a fila | T0 §2 |
| S1.3 | Casos 1 fila (oct-1 ASE3 R23 `286.711.907,53`; jul-2 ASE3 `−16.458,46` si se confirma), 2 filas, 3+ filas (ASE4) | T0 §3 |
| S1.4 | Inventario R2/R4 (conteos por label, read-only): ¿hay 1-fila-empresa en R2? ¿alguna multi-fila en R4? → veredicto D6 | T0 §4 |
| S1.5 | Gate de identidad: re-resolución por sección == resolución por ocurrencia (exacto) en TODO lo verde actual | T0 §5; un mismatch = STOP |
| S1.6 | Gate de tipabilidad: toda fila-empresa tipable por `SeccionPorCierre`; si alguna no → **STOP y reporte** (no inventar) | T0 §6 |
| S1.7 | Arbitraje `TotalMedio`-con-2-secciones: valores del manual-oráculo (`REMUNERACION_*` F222/F362/F368 y sus visibles) en jul-2/ago-2 para ambas hipótesis (`Medio=0` vs `Medio=Menor`) | T0 §7; fija la lista §9 |

### H1 (implementación contract-first, traza a §6–§7)

| ID | Tarea | Traza |
|---|---|---|
| H1.1 | `ResolvedorOperandosEmpresaR1` puro + unit tests (ordinales, ausencias→ausente, empate→señal STOP) | D1/D2/D3, §7 |
| H1.2 | Swap en `ExtraerCeldasR1`: `(fila,indice)` + `SeccionPorCierre` reutilizando construcción `FilaEspejoR1` + resolvedor; ausente→0; retirar `OrderByDescending(\|F\|)` del camino | A1/A2, §7 |
| H1.3 | Doc-comments `FuenteR1` reescritos (semántica D2) | D2 |
| H1.4 | Reescribir `S2_1_ParcialConLabelPresente` a aserciones de valor (baja+alta declarada §9) + tests D5 (2 REALES, desigualdad explícita) | A2/D5, §9 |
| H1.5 | `git diff --check`, CRLF, SHAs, grep de cierre (§8) | §8 |

### H2 (regresión y evidencia)

Suite completa → `TestResults/p52-h2-suite.trx` + `TestResults/p52-h2-setdiff.txt` contra `p51-h2-suite.trx`. PASS criterios §8-iii. Code-review independiente (subagente `code-reviewer`, este plan como spec): **APROBADO** requerido.

### H3 (docs)

Párrafo de doctrina en `.opencode/project-context.md` (mapeo empresa-R1 por sección, tabla D2 resumida, remisión al T0); este plan queda como fuente de verdad. Sin cambios al Instructivo (la huella `FuentesAusentes` se reutiliza, no hay huella nueva).

### 9. Lista cerrada de altas/bajas esperadas (H2 la verifica exacta)

- **Altas**: fixture T0 (~N tests S1.x) · unit tests resolvedor (~M) · 2 no-duplicación D5 · 1 reescritura S2_1 (cuenta como baja+alta).
- **Bajas**: `S2_1_ParcialConLabelPresente_SigueSiendoFailFast` en su forma throw (reemplazada, no eliminada: el método permanece con aserciones de valor).
- **Cambios de valor esperados (Δ congelados por S1.7)**: `TotalMedio` (F222/F362/F368) de 2-secciones → 0; oct-1 ASE3 (throw → valores D3-oportuno); jul-2 ASE3 (throw/duplicado → D3-primera, si S1.1 confirma 1 fila). Cualquier otro Δ, alta o baja = **STOP**.

---

## 11. Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R1 | Jul-2 ASE3 u otro verde actual cambia de valor (D3/Medio→0) más allá de la lista §9 | S1.5 identidad exacta; si diverge → STOP + escala, nunca re-congelar a ciegas |
| R2 | 2 filas-empresa en la MISMA sección (empate sin regla aprobada) | STOP por diseño (H1.1 lo señala, no lo resuelve) |
| R3 | Filas-empresa con `E` distinto dentro de la misma sección (concepto no-Total) | S1.1 registra `E`; si aparece → STOP (el operando exige concepto `Total`) |
| R4 | ASE4/ASE5 con 3 filas donde magnitud≠ordinal | S1.2/S1.5; mismatch → STOP |
| R5 | R2 confirma casos 1-fila reales (duplicado `Total==SubsCont` en verde) | Se declara como residuo con dueño (D6); no expande este plan |
| R6 | Tentar "arreglar" columnas 5/11 por encabezado al tocar el reader | Prohibido por §4.5; el grep de cierre lo vigila |

## 12. Decisiones abiertas para el Ingeniero (el plan NO está aprobado)

1. **Ratificar la lista §9 tras el T0**: en particular los Δ de `TotalMedio→0` (F222/F362/F368) arbitrados por el oráculo manual (S1.7). Si el manual contradice `Medio=0`, el Ingeniero decide entre (a) aceptar la divergencia declarada con dueño, o (b) STOP del plan.
2. **D3-julio-2**: si S1.1 confirma 1 sola fila en jul-2 ASE3, el cambio `Main: −16.458,46 → 0` altera un valor históricamente verde del single-ASE. ¿Se aprueba como corrección (la sección manda) o se declara y se acota?
3. **R2 (D6)**: si S1.4 confirma el defecto gemelo con caso real, ¿se autoriza plan hijo con su T0 (recomendado) o se declara deuda sin plan?
4. **Nada más**: D1/D2/D4/D5 y el método T0→H1→H2 vienen cerrados del encargo; no requieren debate salvo que el T0 dispare un STOP.

**RATIFICACION 2026-10-22 (Ingeniero) — APROBADO 3/3:** (1) si el manual contradice `Medio=0`, se acepta como **divergencia declarada con dueño** (no STOP ciego); (2) julio-2: el `Menor` duplicado **se corrige** (la sección manda; no es valor histórico válido); (3) si el T0 confirma el defecto gemelo en R2 con caso real, se **autoriza plan hijo con su T0**. **H0 HABILITADO.**

---

## 13. Cierre con STOP (2026-10-22)

H0 se detuvo sin tocar código. Evidencia: `Docs/T0-Plan52-Seccion-Operandos-Empresa-R1.md` (299 líneas).

- §9 solo cubría 4 de 17 Δ; 13 no estaban (7 mueven `VisibleR1` hasta −12.823.196.468,26).
- Magnitud ≠ ordinal en verdes (ago-2 ASE2 ENEL/OCCIDENTE).
- Oct-1 ASE5: las 3 empresas en 1 fila de la primera sección; `Main=0` borra `160.186.886,97` de ENERBIT.
- D1: `SeccionPorCierre` vs `PorFronteraArriba` difieren en etiqueta en 180/180.
- El manual confirma la resolución por sección en 4 casos (el código actual es el que está mal).
- Insumos de `3-Ciudad Limpia` revertidos a HEAD por autorización del Ingeniero (2026-10-22); estructura y valores clave intactos.

**Decisión del Ingeniero: Ruta B - rediseño con contrato de sección por operando (Plan 53).**

> **SUPERSEDIDO (2026-10-22):** la Ruta B se implementó en
> `plans/53 - contrato-seccion-por-operando-r1.md` (APROBADO; H0 verde con 2 desviaciones declaradas,
> H1 y H2 cerrados: set-diff `29 altas / 0 bajas / 0 cambios` contra `p51-h2-suite.trx`). El camino
> ordinal queda **prohibido y retirado**: `ExtraerCeldasR1` resuelve por la sección declarada de cada
> celda (`(Celda, FuenteR1, SeccionEsperadaR1)` + `ResolvedorOperandosEmpresaR1`) y el
> `OrderByDescending(|F|)` desapareció de ese camino. Evidencia: `Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md`.
