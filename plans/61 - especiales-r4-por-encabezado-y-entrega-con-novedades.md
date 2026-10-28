---
titulo: Plan 61 - Especiales R4 por encabezado, entrega con novedades y prevencion sistemica
estado: APROBADO 2026-10-28 por el Ingeniero — Clarification Gate §6 ratificado (8/8 con la recomendación senior); H0 habilitado
fecha: 2026-10-28
numero: 61
planes_ocupados: 01..60 ocupados; este es el 61
padre: brief del orquestador 2026-10-28 (defecto P9=0 en R4 + política SIEMPRE-Excel + prevención sistémica)
madre_doctrina: .opencode/project-context.md (plantilla autoridad; values-only; guard anti-fórmula; tolerancia ±0,5; prohibido <v> stale como oráculo; 0 warnings; sin valores quemados por período/ASE/canal; nunca sobrescribir fórmulas) + plans/59 (derivación por etiqueta, candado anti-quemado) + plans/60 (regla pura sin quemados, traza de auditoría, oráculo por etiqueta) + plans/55 (conciliación por empresa, anti-tautología) + plans/42/43 (divergencia declarada, salida "sin validar") + plans/50 (fuentes ausentes, etiquetas no reconocidas)
caso_rojo: dos períodos, dos gates distintos, un solo defecto — `MapearR4` con `P9 = 0m` hardcodeado. 2026092/Q2 ASE4: fuente R4 fila A23="Total"(B vacío) D23=-178.619.709,53, P23(SERVICIO ESPECIALES)=-49.900,44; lo detecta el oráculo R10 (calculado 7.675.312.602 vs R10 7.675.362.502, Δ 49.900,44). 2026091/Q1 ASE2: fuente R4 fila A11="Total"(B vacío) D11=-8.739.602,37, P11=-26.500,00; lo detecta `ValidarSigmaEmpresasPorAse` ("Σ visibles de empresas en R4 (-8.713.102,37) no coincide con el visible de bloque (-8.739.602,37)"). La app calculó 14.147.331.255,39; el administrativo D105=14.147.357.755,39 y R10 DetRetri!D10=14.147.357.755.
gate_entrada:
  plantilla_v4: Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx — VIGENTE. Huella semántica (helper canónico `HuellaSemanticaPlantilla`) a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55 (corregida en H0: la constante canónica manda; el `524d` original era un typo del plan); SHA de archivo informativo. Huella distinta ⇒ STOP.
  suite_base: la vigente al abrir H0 (trx de entrada citado en H0; 0 Skip; set-diff: rojo nuevo fuera de lista = STOP)
  sin_commits: NUNCA git commit / git push; NUNCA git checkout / git reset
  excel_cerrado: Excel debe estar cerrado durante las corridas
---

# Plan 61 — Especiales R4 por encabezado, entrega con novedades y prevención sistémica

> SDD completo PROPOSE → DESIGN → SPEC → TASKS → GATES. Fuente única de verdad para la implementación una vez aprobado ESTE documento. No se implementa sin aprobación del Ingeniero (punto de bloqueo: sin su OK no se abre H0; estado inicial PROPUESTO).
> Lectura obligatoria previa: `plans/59` (derivación por etiqueta, candado anti-quemado, H0 read-only) + `plans/60` (regla pura, coordenada de auditoría, oráculo por etiqueta) + `plans/55` (conciliación por empresa, anti-tautología) + `plans/42` (divergencia declarada, reporte durable, código 6) + `plans/43` (salida "sin validar", promoción por `File.Move`) + `.opencode/project-context.md` (doctrina madre).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/Cli/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español neutro; CRLF sin mezclar.
> Visual Design Intent: mínimo y administrativo (H4): WinForms muestra la ruta del TXT de novedades con botón para abrirlo; sin rediseño, sin premium, sin jerga técnica en lo visible al usuario.

---

## §0. Evidencia citada (la cita el brief; el H0 la re-mide — nada de aquí es oráculo)

### §0.1 El defecto (verificado en código el 2026-10-28; los MONTOS los re-mide H0)

| Dato | Valor citado por el brief (a re-medir) | Verificación en código |
|---|---|---|
| Punto de defecto | `ExcelDataReaderWorkbookLeafInputReader.MapearR4` (~línea 1392): `var p9 = 0m;` con comentario "P = 0 (sin análogo en la fuente, patrón histórico)" | ✅ Existe textual: líneas 1389-1416. `D9` se lee de la columna D de la fila `Total` con B vacío; `P9` es literal `0m` |
| Modelo que sí lo contempla | `WorkbookLeafInputsR4.TotalReversionEsperada => D9 - P9` | ✅ `Remuneracion.Core/Models/WorkbookLeafInputsR4.cs:27`. Con P9=0, P9 se anula por construcción |
| Agregado que también ignora Especiales | `ExcelDataReaderRecaudoReader.LeerR4` (~líneas 181-196): `TotalReversiones` toma solo columna D | ✅ Líneas 170-196: busca fila A=="Total" con B vacío y lee la columna 4 (índice 3). Sin resolución de Especiales |
| Gate que compara leaf vs agregado | `WorkbookLeafCoherence` (líneas 42-44 y 201-202) | ✅ Línea 42-44 (`R4.TotalReversionEsperada vs ReversionR4.TotalReversiones`) y línea 201-202 (vs `ConsolidadoAse.ReversionR4`). Ambos lados ignoran Especiales hoy |
| Asimetría probada | `ResolvedorConciliacionEmpresa` líneas 301-303 YA aplica `VisibleR4 = Σ(Total − Especiales)` por empresa | ✅ Líneas 301-303 + `EncabezadosEspeciales = ["SERVICIO ESPECIALES", "Especiales"]` (línea 50) + `IndicesComponentes` (líneas 555-579) que resuelve ambas columnas por encabezado. El visible por EMPRESA resta Especiales; el visible de BLOQUE no |
| Caso 2026092 ASE4 | R4 D23=-178.619.709,53, P23=-49.900,44; oráculo R10 7.675.362.502 vs calculado 7.675.312.602; administrativo `CONSOLIDADO!D107=7.675.362.501,96`, `D69 = 'Reversion Pagos R4'!D316 = D239 − P239` | ⚠️ Montos y celdas CITADOS, no medidos por este plan. H0 los re-mide en bytes (T0.1). La forma `D239 − P239` del administrativo confirma que la plantilla espera AMBOS operandos |
| Caso 2026091 ASE2 | R4 D11=-8.739.602,37, P11=-26.500,00; gate Σ-empresas (línea ~886); administrativo D67=-8.713.102,37, D105=14.147.357.755,39; R10 DetRetri!D10=14.147.357.755; app 14.147.331.255,39 | ⚠️ Montos CITADOS. H0 los re-mide (T0.2). Gate verificado: `ValidadorBasico.cs:884-886` emite el mensaje citado; el gate corre solo en Q1 (`:866`, residuo `R-SIGMA-Q2` declarado en `:853-865`) |
| Columna intermitente | En 2026092 el ASE 2 trae en la letra P `Ajuste Recargo Aseo` (la plantilla lo ubica en Q) | ⚠️ Citado. H0 lo confirma o corrige (T0.3). Es la prueba de que la columna es INTERMITENTE por ASE y por mes: regla por ENCABEZADO, 0 si el encabezado no existe, nunca por índice |
| Corpus | Barrido de 30 archivos R4 en `Docs/`: solo 2 tienen la columna con valor | ⚠️ Citado con inconsistencia interna (ver §8-C4: el brief dice "30" y "solo 1 de 25" en el mismo texto). H0 hace el barrido propio con conteo cerrado (T0.4) |

### §0.2 Por qué no se detectó (análisis de causa sistémica — el plan lo convierte en candados H5)

| # | Causa citada | Estado verificado | Candado que la cierra (H5) |
|---|---|---|---|
| (a) | Corpus: la columna casi nunca aparece | ⚠️ A re-contar en H0 (T0.4) | Sonda de forma H5.1: todo encabezado no mapeado ⇒ hallazgo auditable |
| (b) | La suite lo bendijo: 3 fixtures con `P9 = 0m` a mano + 1 assert tautológico | ✅ Verificado: `AjustesSfTTests.cs:312`, `ConciliacionEmpresaSigmaPlan55Tests.cs:319`, `ConciliacionEmpresaTests.cs:152` con `P9 = 0m`; `DetalleR2R4Tests.cs:134`: `Assert.Equal(leaf.R4.D9 - leaf.R4.P9, leaf.R4.TotalReversionEsperada)` (identidad de la propia fórmula) | Doctrina de fixtures H5.4 + inventario ampliado (ver §8-C3: hay DOS `P9 = 0m` más no citados) |
| (c) | Un T0 documentó "D − P, con P = 0" como especificación | ✅ Verificado con corrección de nombre (ver §8-C1): `ConciliacionEmpresaCierreBloqueT0Tests.cs:682` (`MapearR4 (D − P, con P = 0)`), más `:395` y `:821` con el mismo supuesto | Auditoría de supuestos H5.3: barrido textual + clasificación + candado anti-literal |
| (d) | El gate de coherencia leaf-vs-agregado lee el mismo archivo con la misma regla ⇒ ciego por construcción | ✅ Verificado: `WorkbookLeafCoherence.cs:42-44` compara dos valores que hoy ignoran Especiales | H1 corrige AMBOS lados en el mismo cambio (S2.1); H5.2 exige que el oráculo derive de la plantilla |
| (e) | El gate Σ-empresas solo corre en Q1; en Q2 la única malla fue el R10 | ✅ Verificado: `ValidadorBasico.cs:866` (`if (leaf.Periodo.NumeroQuincena != 1)` + residuo `R-SIGMA-Q2` en `:853-865`) | Corpus por período H5.5: todo período nuevo entra con su par R10 + Administrativo |

### §0.3 Restricción real del pipeline (verificada en código el 2026-10-28)

`ProcesadorPeriodo.cs`: `_validador.Validar(resultado, leafs)` línea 374 (fail-fast multi-ASE) → `ValidarDetRetriContraR10` línea 389 (fail-fast R10, AMBAS quincenas) → `_workbookLeafWriter.GenerarWorkbook(... rutaSinValidar ...)` línea 402 → promoción `rutaSinValidar → solicitud.RutaSalida` con `File.Move` líneas 539-545. Convención `NombresSalida` (`Remuneracion.Core/Constants/NombresSalida.cs`): sufijo `" (sin validar)"` + `RutaSinValidar(ruta)` (Plan 43 O-4). Reporte durable previo: `IReporteDivergenciasDeclaradas` + `ReporteDivergenciasDeclaradasWriter` emite `<nombre> - Divergencias declaradas.txt` con texto de NEGOCIO (prohibida jerga). Contrato de salida CLI: `CodigosSalida` 0..6 (`DivergenciaDeclaradaReconocida = 6`, Plan 42). WinForms ya opera en dos pasadas con `FormReconocimientoDivergencia` y `UltimoCodigoSalida` (`Form1.cs:486-612`); nunca emite `Environment.Exit`.

**Consecuencia de diseño (importante):** el workbook "sin validar" YA se escribe antes del gate 2.7 y YA queda en disco si la 2.7 falla (`SalidaSinValidarO4Tests`, `AnclasDerivadasH2Plan57Tests.cs:462`). Lo que el Frente 3 cambia no es "escribir o no escribir" sino: (1) mover los gates de DATOS a DESPUÉS de la escritura para que el Excel exista siempre que sea escribible; (2) emitir el TXT de novedades; (3) estado y código de salida inequívocos; (4) prohibición de falso verde (nunca nombre limpio, nunca `OK`).

---

## §1. PROPOSE — Problema / Objetivo

### 1.1 Intent

Que la app lea `SERVICIO ESPECIALES` del R4 **por encabezado** (0 si el encabezado no existe) en TODOS los caminos que derivan el mismo número (agregado, leaf por empresa, leaf de bloque, gates de coherencia); que **ninguna divergencia de datos deje al usuario sin Excel** —toda corrida con workbook escribible entrega el Excel marcado `(sin validar)` más un TXT de novedades en lenguaje de negocio—; y que la clase completa de defecto ("columna intermitente ignorada en silencio") quede imposible por construcción mediante sonda de forma, candado derivado de la plantilla y doctrina de fixtures. Sin quemar ASE/período/columna-letra en ningún camino nuevo.

### 1.2 In Scope (H0 + H1 + H2 + H3 + H4 + H5, prohibido ampliar sin nuevo OK del Ingeniero)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **H0/T0** | Verificación read-only: re-medición en bytes de los dos casos, inventario exacto de gates y de supuestos de ausencia, barrido de corpus con conteo cerrado | Fixture verde + doc de evidencia; puerta H0→H1 |
| **H1** | Fix de lectores R4 (agregado `LeerR4` + leaf `MapearR4` + coherencia) por encabezado | CA-C1 en corridas reales 2026091/2026092 |
| **H2** | Fixtures y tests de regresión (casos 2026091-ASE2 y 2026092-ASE4 + `P9 = 0m` reemplazados + regresión "sin columna → 0") | CA-C2 + suite verde con set-diff |
| **H3** | Política de entrega con novedades (orden escritura/validación, modelo de resultado, clasificación de gates) | CA-C3 en corridas reales con divergencia |
| **H4** | Reporte TXT de novedades + superficies CLI/WinForms + códigos de salida | CA-C4 en corridas reales + revisión UX del Ingeniero |
| **H5** | Prevención sistémica (sonda de encabezados, candado derivado de plantilla, doctrina de fixtures, auditoría de supuestos) | CA-C5 + candados verdes |
| **Gate** | Corridas CLI 2026091, 2026092 (casos), 2026071/2026072/2026082/2026101 (no-regresión) + suite completa con set-diff | CA-C6 |

### 1.3 Non-goals / out (vinculante — ver §9)

- No tocar la plantilla (huella v4 intacta; cero `<f>`, cero `InsertRow`/`DeleteRow`).
- No tocar la lógica de cálculo (`CalculoRemuneracion`), ni CONSOLIDADO, ni `DetRetri`/`DetValiRetri`, ni `VALIDACION_TOTAL`.
- No reabrir Plan 59 (delimitación por cierres), Plan 60 (reclasificación Deb/Cred), Plan 55/56, ni Plan 58-P2.
- No re-congelar letras de columna, filas, ASE, períodos ni quincenas bajo ninguna forma (ni mapa, ni allow-list, ni "solo este ASE").
- Commits, docs de usuario. NUNCA `git commit` / `git push`. NUNCA `git checkout` / `git reset`.

### 1.4 Riesgos con mitigación (resumen; detalle en §7)

| ID | Riesgo | Mitigación |
|---|---|---|
| R-LETRA | Resolver Especiales por letra fija (P) en vez de por encabezado — M-B del §2 | D-H1 (encabezado, 0 si ausente) + U-regresión "sin columna → 0" + candado H5.2 |
| R-UN-LADO | Corregir el leaf y olvidar el agregado (o viceversa) — el gate salta | S2.1: ambos lados en el mismo cambio + test de coherencia con Especiales ≠ 0 |
| R-FALSO-VERDE | Entregar con novedades reportando `OK` o con nombre limpio | D-H3 (estado inequívoco + código nuevo + nombre siempre marcado) + CA-C3 + test que lo fija |
| R-REPORTE-JERGA | El TXT habla en `ERR-*`, `<f>` o nombres de clase y el administrativo no lo entiende | D-H4 (lenguaje de negocio, precedente del reporte de declaradas) + revisión del Ingeniero |
| R-ORACULO-FILA | Comparar contra el administrativo/R10 por fila (geometrías distintas) | Prohibido normativo (§0.6 del Plan 60, vigente); todo gate por etiqueta/nivel |
| R-ORACULO-STALE | Usar `<v>` caché de celdas con fórmula del oráculo como verdad | G-WORKBOOK: texto-`<f>` + literales; `<v>` de fórmula nunca es oráculo |
| R-REGRESION | El fix mueve períodos históricos que hoy cierran | CA-C6: 2026071/2026072/2026082/2026101 byte-idénticos en R4 + suite set-diff |
| R-DOBLE-REPORTE | Dos TXT que se contradicen (declaradas vs novedades) | D-H4: un solo reporte con secciones; compatibilidad del nombre anterior fijada por test |

---

## §2. DESIGN

### D-H0. Método T0 por bytes (read-only bloqueante, cero código de producción)

Forma normativa: BCL puro (`ZipArchive` + `X.Linq`) para plantilla y administrativos + ExcelDataReader solo-lectura sobre copias para fuentes R4; rutas explícitas a `Docs/Prueba Septiembre 1/Insumos`, `Docs/Prueba Septiembre-2/Insumos`, sus R10 y sus administrativos. Prohibido `Insumos.Raiz()` en lo nuevo. Cero código de producción en H0: solo fixture `*T0Plan61*` + doc `Docs/T0-Plan61-Especiales-R4.md`.

Seis frentes: (a) re-medir D23/P23 (2026092 ASE4) y D11/P11 (2026091 ASE2) en bytes: fila, letra de columna física, encabezado exacto (incluido el `Ajuste Recargo Aseo` en P del ASE2-2026092); (b) re-medir `D69 = D239 − P239` (o la forma real) en el administrativo y `D67/D105` + `DetRetri!D10` en el otro; (c) barrido de corpus: conteo cerrado de archivos R4 en `Docs/` con presencia/ausencia del encabezado y valores; (d) inventario exacto de gates que tocan R4 (Validar Σ-empresas, coherencia leaf-vs-agregado, coherencia vs consolidado, ValidarDetRetriContraR10, 2.7) con su posición antes/después del writer; (e) inventario exacto de supuestos de ausencia (`= 0m`, "sin análogo", "patrón histórico", "no tiene análogo") en lectores, con archivo:línea; (f) inventario de tests que asumen ausencia de salida o nombre limpio (ProcesadorPeriodoTests, SalidaSinValidarO4Tests, AnclasDerivadasH2, VerificadorRecaudo T19, nombre del reporte de declaradas).

**Puerta H0→H1:** seis frentes cerrados + fixture verde. Si un número del §0 difiere ⇒ STOP (se re-mide, no se "ajusta" el plan en silencio). Las contradicciones §8-C1..C6 quedan resueltas o escaladas aquí.

| Candidata | Veredicto | Motivo |
|---|---|---|
| **M-A: T0 de seis frentes + fix por encabezado + entrega-con-novedades + candados sistémicos** | **Elegida** | Única que cierra el defecto Y la clase (Frentes 1+2+3 del brief) |
| M-B: leer Especiales por letra fija (P, o P/Q según ASE) | Descartada | Re-fotografía la intermitencia: el próximo mes la columna se mueve y rompe en silencio o con falso verde |
| M-C: fix solo del leaf, dejando el agregado en D | Descartada | El gate de coherencia (ciego hoy, honesto mañana) saltaría: R-UN-LADO |
| M-D: fail-fast honesto sin Excel (no entregar nada ante divergencia) | Descartada | Contradice el requisito explícito del Ingeniero (Frente 3): sin Excel "nos quedamos sin herramientas" |

### D-H1. Fix de lectores R4 por encabezado (decisión)

1. **Regla única:** `Especiales(R4) = valor en la fila de cierre (A="Total", B vacío) bajo la columna cuyo encabezado normalizado sea `SERVICIO ESPECIALES` (alias admitido: `Especiales`); 0 si ningún encabezado lo es.` Normalización = `Trim` + `OrdinalIgnoreCase` (misma convención que `EsEncabezadoEspeciales` y que `R2-ServEspK`). La letra física (P hoy, otra mañana) NUNCA se nombra en producción.
2. **Reutilizar, no reinventar:** `ResolvedorConciliacionEmpresa.EncabezadosEspeciales` + `IndicesComponentes` ya resuelven `(Total, Especiales)` por encabezado para R4 por empresa. `MapearR4` y `LeerR4` consumen la misma tabla de encabezados (mismo `string[]`, no una copia divergente). Si el consumo directo no es posible por visibilidad, se expone un accesorio interno puro —nunca se duplica la lista de alias.
3. **Tres toques, un solo cambio lógico (R-UN-LADO):**
   - `MapearR4`: `P9 = Especiales(filaTotal)`; `ValoresPorRol[GranTotal(EspecialesR4)] = P9` (el writer ya escribe `[(HojaR4,"P9")] = leaf.R4.P9` en `OpenXmlPlantillaWriter.cs:1009`: con P9≠0 fluye solo).
   - `LeerR4` (`ExcelDataReaderRecaudoReader.cs:170-196`): resuelve la columna Especiales por encabezado y expone el valor. El modelo `ReversionR4` se extiende **aditivamente** (p. ej. `EspecialesReversion`, default `0m`): `TotalReversiones` sigue siendo el D observado ("observación fiel de la fuente", doctrina `DetalleR2AseInputs`), y la coherencia compara `leaf.TotalReversionEsperada (D9−P9)` vs `(r4.TotalReversiones − r4.EspecialesReversion)`. H0 verifica la forma exacta del modelo antes de fijar el nombre.
   - `WorkbookLeafCoherence` (`:42-44` y `:190-205`): ambos lados con Especiales. Ningún otro gate cambia su aritmética: Σ-empresas ya resta (líneas 301-303), `ValidarDetRetriContraR10` y 2.7 leen el workbook.
4. **Comentarios obsoletos:** los textos "P = 0 (sin análogo en la fuente, patrón histórico)" (`:1389-1390`), "(P = 0 en la fuente)" (`ResolvedorConciliacionEmpresa.cs:301`) y sus gemelos del T0 (`ConciliacionEmpresaCierreBloqueT0Tests.cs:395,682,821`) se actualizan en el mismo cambio: un comentario que bendice el supuesto es parte del defecto (causa §0.2-c).

### D-H3. Política de entrega con novedades (decisión — valida o mejora la propuesta del brief)

**Criterio general (se adopta el propuesto, formalizado):** degrada a *novedad* toda **divergencia de DATOS con workbook escribible** (el número calculado no coincide con el esperado pero el artefacto existe y es revisable); sigue siendo *fallo duro* toda **imposibilidad de construir el artefacto** (no hay nada revisable que entregar).

| Clase | Gates | Efecto |
|---|---|---|
| **Novedad** (entrega `(sin validar)` + TXT + código nuevo, NUNCA nombre limpio, NUNCA `OK`) | Σ visibles-empresa vs bloque R1/R2/R4 (`ValidadorBasico:846-886` + `WorkbookLeafCoherence:97-122`); DetRetri calculado vs R10 (`ValidarDetRetriContraR10`); validación cruzada 2.7 no declarada (`ProcesadorPeriodo:453-459`); divergencia declarada sin reconocimiento (flujo Plan 42 intacto, ahora también con Excel + TXT) | `EstadoEntrega.EntregadaConNovedades`, colección de novedades tipadas, `RutaReporteNovedades`, exit nuevo |
| **Fallo duro** (sin promoción, códigos intactos, comportamiento actual) | Preflight de insumos faltantes; fuente ilegible/formato (`ERR-FORMATO-FUENTE`); invariantes estructurales del writer (`ERR-PLANTILLA`: plantilla alterada, hoja/destino inexistente, celda-fórmula donde va valor); `ERR-ESCRITURA`/`ERR-INESPERADO` | Throw antes de promover; el `sin validar` ya escrito queda como hoy (evidencia, no entrega); sin TXT de novedades |

**Decisiones que este diseño toma (abiertas a ratificación en §6):**
- (i) La 2.7 no declarada (`VALIDACION_TOTAL` FALSE en el workbook) degrada a novedad: el workbook dice FALSE en su propia fórmula y ESO es lo que el TXT explica (qué celda, qué Δ, contra qué administrativo). Negar el Excel no hace el FALSE menos real; entregarlo lo hace revisable.
- (ii) Orden nuevo en `ProcesadorPeriodo`: preflight (duro) → lecturas (duro si ilegible) → cálculo → **escritura a `rutaSinValidar`** → gates de datos (novedad) → 2.7 (novedad o declarada) → promoción solo si cero novedades. El `sin validar` existe siempre que la escritura fue posible.
- (iii) **Prohibición de falso verde** (normativa): con novedades, el resultado NUNCA porta `Ok`, el CLI NUNCA imprime `RESULTADO OK`, WinForms NUNCA muestra verde, y el archivo NUNCA toma el nombre limpio. Huella durable doble: TXT + log con RunId.

**Nuevo resultado de dominio (aditivo, compatible):**

```csharp
// Propuesta (nombres exactos a fijar en implementación contra la convención del directorio):
public enum EstadoEntrega { Certificada, EntregadaConNovedades }
public sealed record NovedadEntrega(
    string Reporte, string Hoja, int AseId, string ConceptoOCelda,
    decimal ValorObservado, decimal ValorEsperado, decimal Delta,
    string Dueno, string QueRevisar);   // lenguaje de negocio, sin ERR-*, sin <f>, sin clases
// ResultadoProcesoPeriodo += EstadoEntrega (default Certificada),
//   Novedades (default vacía), RutaReporteNovedades (default vacía).
// Invariante: Novedades no vacía ⇒ Estado == EntregadaConNovedades,
//   RutaSalida termina en " (sin validar)", y el código de salida ≠ 0.
```

### D-H4. Reporte TXT + superficies (decisión)

- **Un solo reporte con secciones** (recomendación del brief, se adopta): el TXT de novedades absorbe como sección el contenido del ` - Divergencias declaradas.txt` cuando la corrida también cierra declaradas reconocidas. **Compatibilidad:** el nombre anterior se conserva cuando SOLO hay declaradas (tests y manuales lo citan — H0 inventaría qué tests dependen del nombre exacto, T0.6); con novedades de datos el archivo es `<nombre> - Novedades de la corrida.txt`. Si H0 demuestra que ningún test depende del nombre, el Ingeniero puede optar por nombre único siempre (§6, decisión D-TXT).
- **Estructura:** (1) encabezado con veredicto inequívoco `ENTREGADO CON NOVEDADES — NO CERTIFICADO`; (2) datos de corrida (período, quincena, runId, usuario, canal, fecha/hora, ruta del Excel); (3) una ficha por novedad (reporte/hoja/ASE/concepto-o-celda, observado, esperado, Δ, dueño, qué revisar); (4) guía "cómo revisar" (hoja/celda del Excel vs Administrativo/R10, por etiqueta, nunca por fila); (5) sección de declaradas si aplica. UTF-8 sin BOM (precedente Plan 42).
- **Determinismo:** re-ejecutar la misma corrida reescribe/actualiza el mismo TXT (misma entrada → misma ruta, como `NombresSalida`).
- **CLI (`EjecutorCli`):** mensaje `RESULTADO CON NOVEDADES codigo=<nuevo> salida=<ruta-sin-validar> reporte=<txt> runId=<guid>` en stdout (grepable, precedente `RESULTADO OK`/`RESULTADO ERROR`); detalle en stderr/log. **Código nuevo 7** (`CodigosSalida.NovedadesEntrega = 7`, nombre exacto a ratificar): distinto de 0 y de 6 por diseño; contrato 0..6 intacto. `CatalogoErrores` NO lo mapea desde `CodigoError` (no es un error: es un estado de entrega) —el CLI lo retorna directo; H0 verifica que ningún test asuma `<= 6`.
- **WinForms (`Form1`):** con novedades, `UltimoCodigoSalida = 7`, status ámbar (nunca verde), log con la ruta del TXT y **botón para abrirlo** (precedente: el cierre con declaradas ya muestra `Reporte: {ruta}` en `:603-606`; se extiende el mismo patrón, no un flujo nuevo).

### D-H5. Prevención sistémica (decisión — se adopta y precisa la recomendación del brief)

1. **Sonda de forma de insumos (H5.1):** por cada archivo fuente del período, enumerar encabezados observados y contrastarlos contra el conjunto que el código consume por (reporte, hoja). Todo encabezado no mapeado ⇒ hallazgo `ASE n · Reporte · Hoja · Encabezado` (formato heredado de `EtiquetasNoReconocidas`). Nivel: **novedad auditable** por defecto; **fail-fast** solo si coexiste con `FuentesAusentes` del MISMO reporte (precedente Plan 50 A3: posible fuente presente no reconocida). Vive en el reader (observación fiel) y viaja en `ResultadoProcesoPeriodo.EtiquetasNoReconocidas` (campo ya existente —se reutiliza, no se duplica).
2. **Candado de forma derivado de la plantilla (H5.2):** test que, por cada celda de CONSOLIDADO cuya fórmula sea `<hoja>!X<fila> − <hoja>!Y<fila>` (forma Total − Especiales), exija que el lector exponga AMBOS operandos resueltos por encabezado. **Derivado de la v4 en disco** (se lee la plantilla, no una lista congelada): si la plantilla gana otra resta, el candado exige su operando sin cambiar código de test.
3. **Prohibición de supuestos de ausencia (H5.3):** T0 textual sobre `= 0m`, "sin análogo", "patrón histórico" en lectores (inventario T0.5); cada supuesto se clasifica (legítimo-documentado vs defecto-latente) y lo latente se corrige; candado que impide re-introducir literales 0 en valores derivados de fuente (allow-list explícita con motivo para los legítimos, p. ej. "columna ausente → 0" con su test de regresión).
4. **Doctrina de fixtures (H5.4):** prohibido inventar valores de fuente (`P9 = 0m`, `E26 = 0m` a mano); los agregados vienen de fuente real o se declaran explícitamente sintéticos (nombre del fixture + comentario); prohibidos asserts tautológicos (identidad de la propia fórmula): todo assert de agregado exige un valor medido independiente (fuente, oráculo o cálculo ortogonal).
5. **Corpus por período (H5.5):** todo período nuevo entra con su par R10 + Administrativo como oráculos; el gate R10 (`ValidarDetRetriContraR10`, AMBAS quincenas) y el gate Σ-empresas en Q1 son la malla mínima. Sin par de oráculos ⇒ el período entra marcado y H0 del plan que lo abra lo declara.
6. **Auditoría de simetría (H5.6, checklist de cierre):** R2 ya resta Especiales (`ServEspK`); R1 ya resta L; R4 no restaba (este plan). H5 verifica que no exista el mismo hueco en otras columnas/familias (DetalleR2: `Deb/Cred-otros`; R1: canales por encabezado tras el Plan 60; BCE/anticipos) y lo deja en tabla con veredicto por familia.

### Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Lectores leen (por encabezado); coherencia compara; procesador ordena (escribir→validar); reporte narra; sonda observa. Cada pieza una sola razón de cambio |
| **OCP** | ✅ Columna nueva/movida se absorbe por encabezado sin código nuevo; gate nuevo degrada a novedad por la clasificación general (§2-D-H3), no caso por caso |
| **LSP/DIP** | ✅ Contratos `IRecaudoReader`/`IPlantillaWriter`/`IValidador` intactos; `ResultadoProcesoPeriodo` solo aditivo (defaults = comportamiento anterior); cambios detrás de las mismas abstracciones |
| **Best practices** | ✅ Contract-first (H0 congela bytes antes del juicio); fail-fast donde no hay artefacto + novedad donde lo hay; tolerancia ±0,5 solo en comparación; values-only; guard anti-fórmula intacto; CRLF; sin commits |
| **Performance** | ✅ H0 acotado BCL + lectura solo-copias; sonda O(encabezados) por archivo una vez por corrida; cero workbooks extra en runtime; sin Excel en los gates |

---

## §3. SPEC (requisitos numerados + gates no negociables)

### S0 — Gates transversales (toda H)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa trx + **set-diff de rojos vs base de entrada**: altas/bajas una por una con causa citada; **rojo nuevo fuera de lista ⇒ STOP**. 0 `Skip`, 0 tests borrados, 0 tolerancias ampliadas; re-anclajes solo con evidencia.
- G-HUELLA: huella v4 `a7476ee2…8c55` intacta antes/después de cada H (el SHA de archivo es informativo).
- G-FORMULA: `<f>` byte-idéntica en todo el diff (cero `new CellFormula` fuera de `EscribirFormulaPreservando`, cero `InsertRow`/`DeleteRow`). **Cero `<f>` escritas por nosotros.**
- G-QUEMADO: cero literales de letra de columna/fila/celda y cero ramas por ASE/período/quincena en los caminos tocados —incluida la forma escondida en mapas y mensajes— (grep de cierre + candado §5-H5).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale) o la fuente ya leída; **cero workbooks extra en runtime**; **sin Excel en el gate**.
- G-TRAZA: cada novedad y cada error nuevo nombra reporte + hoja + ASE + concepto/celda (+ observado vs esperado + Δ donde aplique). Prohibida jerga en el TXT (`ERR-*`, `<f>`, nombres de clase).
- G-PROC: cero commits; cero `git checkout` / `git reset`; CRLF sin mezclar en lo tocado; Excel cerrado durante las corridas.
- G-H0-READONLY: H0 no contiene código de producción (solo fixture `*T0Plan61*` + doc `Docs/T0-Plan61-Especiales-R4.md`); el diff de `Remuneracion.Core/` + `Remuneracion.Infrastructure/` + `Remuneracion.Cli/` + `Remuneracion.WinForms/` en H0 es vacío salvo esos tests.

### CA — Criterios de aceptación (vinculantes; todo número viaja con su derivación fuente × etiqueta)

- **CA-C1 (lectores):** `MapearR4` y `LeerR4` resuelven Especiales por encabezado; `WorkbookLeafCoherence` compara ambos lados con Especiales; comentarios "P = 0 / sin análogo / patrón histórico" actualizados (código + T0). Derivación: encabezados medidos en T0.1/T0.2 × `EncabezadosEspeciales` único.
- **CA-C2 (casos):** 2026091 ASE2 DetRetri = **14.147.357.755** y 2026092 ASE4 DetRetri = **7.675.362.502** (o los valores que H0 re-mida si el §0 difiere —el T0 manda), con Excel `(sin validar)` entregado + TXT de novedades. Derivación: fuente R4 × oráculo R10/administrativo por etiqueta.
- **CA-C3 (falso verde imposible):** con novedades, ningún camino reporta `OK`/0/verde/nombre limpio; el TXT existe y el log porta el RunId. Fijado por test negativo (injerto: novedad sin TXT ⇒ rojo; novedad con nombre limpio ⇒ rojo).
- **CA-C4 (superficies):** CLI imprime `RESULTADO CON NOVEDADES` con código 7 y rutas; WinForms muestra status ámbar + ruta del TXT + botón de apertura; `UltimoCodigoSalida = 7`. Revisión UX del Ingeniero.
- **CA-C5 (sistémico):** sonda H5.1 verde (incluido injerto: encabezado nuevo no mapeado ⇒ hallazgo); candado H5.2 derivado de la v4 verde; auditoría H5.3 con cero supuestos latentes; doctrina H5.4 (los `P9 = 0m`/`E26 = 0m` manuales reemplazados + tautología `DetalleR2R4Tests:134` reemplazada); corpus H5.5 con par de oráculos por período; checklist H5.6 con veredicto por familia.
- **CA-C6 (gates):** build 0 warnings; suite con set-diff exacto; 0 regresiones en 2026071/2026072/2026082/2026101 (R4 byte-idéntico donde la columna no existe); huella v4 intacta; review independiente **APROBADO**.

### S1 — Criterios H0/T0 (read-only, cero código de producción)

- S1.1: D23/P23 (2026092 ASE4) y D11/P11 (2026091 ASE2) re-medidos: fila, letra física, encabezado exacto, incluido `Ajuste Recargo Aseo` en P del ASE2-2026092.
- S1.2: forma real de `D69 = D239 − P239` (2026092) y `D67/D105` + `DetRetri!D10` (2026091) re-medida por etiqueta en administrativos/R10.
- S1.3: encabezado exacto de `SERVICIO ESPECIALES` en ambas fuentes (normalización que aplica) + ausencia/presencia por ASE en ambos períodos.
- S1.4: barrido de corpus con conteo cerrado (archivos R4 × presencia del encabezado × valores ≠ 0); resuelve §8-C4.
- S1.5: inventario exacto de supuestos de ausencia (archivo:línea) + inventario de fixtures con valores de fuente inventados (incluidos los no citados en el brief, §8-C3) + inventario de tests que asumen ausencia de salida/nombre limpio/nombre de reporte (T0.6).
- S1.6: lista cerrada de Δ de suite + huella v4 antes/después de H0. **Sin H0 verde no se abre H1.**

### S2 — Criterios H1 (fix por encabezado) + H2 (fixtures y regresión)

- S2.1: `MapearR4` + `LeerR4` + `WorkbookLeafCoherence` (ambos lados) resuelven Especiales por el encabezado único; `TotalReversiones` sigue siendo D observado; comentarios obsoletos actualizados. Grep G-QUEMADO: ninguna letra de columna R4 en producción.
- S2.2: regresión "sin columna → 0" verde (período/ASE sin el encabezado: P9 = 0 por ausencia declarada, no por literal) + caso sintético "columna movida de letra" (misma fuente con el encabezado en otra letra ⇒ mismo P9).
- S2.3: fixtures con `P9 = 0m`/`E26 = 0m` manuales reemplazados por valores de fuente real o marca explícita de sintético; `DetalleR2R4Tests:134` (y gemelo `:161`) reemplazado por assert contra valor medido independiente.
- S2.4: CA-C1 + CA-C2(parcial H1: números cierran) + CA-C6(parcial: 0 regresiones históricas) + 0 warnings + huella intacta.

### S3 — Criterios H3 (entrega con novedades) + H4 (TXT + superficies)

- S3.1: orden escribir→validar en `ProcesadorPeriodo`; `EstadoEntrega` + `NovedadEntrega` + `RutaReporteNovedades` aditivos con defaults = comportamiento anterior; invariante (novedades ⇒ marcado + código ≠ 0) fijada por test.
- S3.2: clasificación Novedad vs Fallo duro implementada según §2-D-H3; cada gate de datos del inventario T0.4 tiene su clase citada; 2.7-no-declarada degrada a novedad con su ficha (celda FALSE + Δ + oráculo).
- S3.3: TXT con las 5 secciones (§2-D-H4), lenguaje de negocio, determinista (re-ejecución ⇒ misma ruta/contenido); compatibilidad del nombre ` - Divergencias declaradas.txt` según veredicto T0.6.
- S3.4: CLI `RESULTADO CON NOVEDADES` + código 7 + rutas; WinForms ámbar + ruta + botón; `UltimoCodigoSalida = 7`; contrato 0..6 intacto.
- S3.5: CA-C2/C3/C4 en corridas reales 2026091 (`--cinco-ase`) y 2026092 + 0 warnings.

### S4 — Criterios H5 (prevención sistémica) + cierre

- S4.1: sonda H5.1 en producción (readers) + hallazgos en `EtiquetasNoReconocidas` con formato `ASE n · Reporte · Hoja · Encabezado`; injerto (encabezado nuevo) ⇒ hallazgo; coexistencia con `FuentesAusentes` ⇒ fail-fast (precedente Plan 50 A3).
- S4.2: candado H5.2 derivado de la v4 en disco (ninguna lista congelada de celdas); si la plantilla gana una resta Total−Especiales, el candado la exige sin editar el test.
- S4.3: auditoría H5.3 cerrada (cero supuestos latentes; legítimos en allow-list con motivo + test); doctrina H5.4 con candado textual (grep en CI de test: `P9 = 0m`/`E26 = 0m` manual ⇒ rojo salvo marca explícita de sintético).
- S4.4: corpus H5.5 (par R10+Administrativo por período, tabla en el plan de cierre) + checklist H5.6 con veredicto por familia.
- S4.5: suite trx + set-diff vs base; code-review independiente (`code-reviewer`, este plan como spec): **APROBADO** requerido; doctrina en `.opencode/project-context.md` (solo lo probado).

---

## §4. TASKS (fases H0..H5 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero código de producción)

Fixture: `Remuneracion.IntegrationTests/EspecialesR4T0Plan61Tests.cs` (BCL ZIP+XML + ExcelDataReader solo-lectura sobre copias; rutas explícitas; prohibido `Insumos.Raiz()` en lo nuevo). Evidencia: `Docs/T0-Plan61-Especiales-R4.md`.

| ID | Tarea | Verifica |
|---|---|---|
| T0.1 | Re-medir 2026092 ASE4 en bytes: D23/P23 (fila, letra, encabezado), `Ajuste Recargo Aseo` en P del ASE2, forma `D69 = D239 − P239` en el administrativo, R10 7.675.362.502 | S1.1 + S1.2(parcial) |
| T0.2 | Re-medir 2026091 ASE2 en bytes: D11/P11, `D67/D105` administrativo, `DetRetri!D10` R10, Δ 26.500,00 | S1.1 + S1.2(parcial) |
| T0.3 | Fijar encabezado exacto + normalización aplicable en ambas fuentes; presencia/ausencia por ASE en ambos períodos | S1.3 |
| T0.4 | Barrido de corpus (conteo cerrado R4 × encabezado × valores) + inventario exacto de gates R4 con posición antes/después del writer | S1.4 |
| T0.5 | Inventario de supuestos de ausencia (archivo:línea) + fixtures con valores inventados (los 3 citados + los omitidos §8-C3) + tautologías (`DetalleR2R4Tests:134,161`) | S1.5(parcial) |
| T0.6 | Inventario de tests que asumen ausencia de salida / nombre limpio / nombre del reporte (`ProcesadorPeriodoTests.GateDeCoherenciaRoto…_NoGeneraArchivo`, `SalidaSinValidarO4Tests`, `AnclasDerivadasH2Plan57Tests:457-463`, `VerificadorRecaudo` T19, tests del nombre ` - Divergencias declaradas.txt`) + verificación "ningún test asume exit ≤ 6" | S1.5(parcial) + S1.6 |
| T0.7 | Lista cerrada de Δ de suite + huella v4 antes/después de H0; resolución o escalado de §8-C1..C6 | S1.6 |

- **Puerta H0→H1:** fixture verde + S1.1-S1.6 cerrados + APROBADO del Ingeniero + D1..D6 del §6 ratificadas donde aplique. Si un número del §0 difiere ⇒ STOP (se re-mide; el T0 manda).

### H1/F1 — Fix de lectores R4 por encabezado (traza a §2-D-H1)

| ID | Tarea | Traza |
|---|---|---|
| H1.1 | `MapearR4`: `P9 = Especiales(filaTotal)` por encabezado único (reutiliza `EncabezadosEspeciales`); `ValoresPorRol[GranTotal(EspecialesR4)] = P9`; actualiza XML doc `:1389-1390` | D-H1.2/1.3/1.4 |
| H1.2 | `LeerR4`: resuelve columna Especiales por encabezado; extiende `ReversionR4` aditivamente (nombre exacto según T0 del modelo); `TotalReversiones` sigue = D observado | D-H1.3 |
| H1.3 | `WorkbookLeafCoherence` (`:42-44`, `:190-205`): ambos lados con Especiales; actualiza comentario obsoleto `ResolvedorConciliacionEmpresa.cs:301` y T0 `:395,682,821` | D-H1.3/1.4 |
| H1.4 | e2e 2026091 + 2026092: CA-C1 (números cierran al centavo contra R10/administrativo por etiqueta) | S2.4 |
| H1.5 | Greps G-QUEMADO (ninguna letra de columna R4 en producción) + huella + 0 warnings | S2.4 |

- **Puerta H1→H2:** S2.1 + CA-C1 + 0w + huella + greps.

### H2/F2 — Fixtures y regresión (traza a §0.2-b y H5.4)

| ID | Tarea | Traza |
|---|---|---|
| H2.1 | Casos 2026091-ASE2 (Δ 26.500,00) y 2026092-ASE4 (Δ 49.900,44) como tests de regresión e2e/leaf (valores del T0, por etiqueta) | S2.4, CA-C2 |
| H2.2 | Regresión "sin columna → 0" (período sin encabezado: P9 = 0 por ausencia) + "columna movida de letra" (mismo valor en otra letra ⇒ mismo P9) | S2.2, anti-M-B |
| H2.3 | Reemplazar `P9 = 0m` manuales (`AjustesSfTTests:312`, `ConciliacionEmpresaSigmaPlan55Tests:319`, `ConciliacionEmpresaTests:152` + omitidos T0.5) y tautologías (`DetalleR2R4Tests:134,161`) por valores medidos o marca explícita de sintético | S2.3, H5.4 |
| H2.4 | Suite trx + set-diff vs base; fuera de lista = STOP; 0 regresiones históricas (2026071/2026072/2026082/2026101) | S2.4, CA-C6 |

- **Puerta H2→H3:** S2.2/S2.3/S2.4 verdes + 0w + huella.

### H3/F3 — Política de entrega con novedades (traza a §2-D-H3)

| ID | Tarea | Traza |
|---|---|---|
| H3.1 | Nuevo resultado: `EstadoEntrega` + `NovedadEntrega` + `RutaReporteNovedades` aditivos en `ResultadoProcesoPeriodo` (defaults = anterior) + invariante en test | D-H3, S3.1 |
| H3.2 | Reorden en `ProcesadorPeriodo`: escribir a `rutaSinValidar` ANTES de los gates de datos; cada gate del inventario T0.4 con su clase (Novedad vs Duro); promoción solo con cero novedades | D-H3, S3.2 |
| H3.3 | Revisar tests del inventario T0.6 preservando intención (no-certificar ≠ no-escribir): `…_NoGeneraArchivo` pasa a exigir "no existe nombre LIMPIO + existe `(sin validar)`"; `SalidaSinValidarO4Tests`/`AnclasDerivadasH2:462` intactos en espíritu; `VerificadorRecaudo` T19: `noCertificado` con archivo `(sin validar)` presente | S3.5, §8-C5 |
| H3.4 | Tests negativos anti-falso-verde (novedad sin TXT ⇒ rojo; novedad con nombre limpio ⇒ rojo; novedad con `OK`/0 ⇒ rojo) | S3.1, CA-C3 |
| H3.5 | e2e 2026091 + 2026092: Excel `(sin validar)` + novedades tipadas + sin promoción | S3.5, CA-C2/C3 |

- **Puerta H3→H4:** S3.1/S3.2/S3.5 verdes + 0w + anti-falso-verde verde.

### H4/F4 — TXT de novedades + superficies (traza a §2-D-H4)

| ID | Tarea | Traza |
|---|---|---|
| H4.1 | Writer del TXT (5 secciones, negocio, UTF-8 sin BOM, determinista); un solo reporte con sección de declaradas; compatibilidad del nombre anterior según T0.6 (decisión D-TXT del §6) | D-H4, S3.3 |
| H4.2 | CLI: `RESULTADO CON NOVEDADES codigo=7 salida=<sin-validar> reporte=<txt> runId=<guid>`; `CodigosSalida.NovedadesEntrega = 7`; contrato 0..6 intacto | D-H4, S3.4 |
| H4.3 | WinForms: status ámbar + ruta del TXT + botón de apertura; `UltimoCodigoSalida = 7`; reutiliza el patrón `:598-606` (no un flujo nuevo) | D-H4, S3.4 |
| H4.4 | e2e ambas quincenas + revisión UX del Ingeniero (lenguaje del TXT y guía "cómo revisar") | S3.4/S3.5, CA-C4 |

- **Puerta H4→H5:** S3.3/S3.4 verdes + UX ratificada.

### H5/F5 — Prevención sistémica + cierre (traza a §2-D-H5)

| ID | Tarea | Traza |
|---|---|---|
| H5.1 | Sonda de encabezados en readers + hallazgos `ASE n · Reporte · Hoja · Encabezado` en `EtiquetasNoReconocidas`; fail-fast solo con `FuentesAusentes` del mismo reporte; test de injerto | D-H5.1, S4.1 |
| H5.2 | Candado derivado de la v4 en disco (CONSOLIDADO `<hoja>!X − <hoja>!Y` ⇒ ambos operandos por encabezado); cero listas congeladas | D-H5.2, S4.2 |
| H5.3 | Auditoría de supuestos T0.5: corrección de latentes + allow-list motivada de legítimos + candado textual | D-H5.3, S4.3 |
| H5.4 | Doctrina de fixtures (cierre de H2.3) + candado anti-`P9 = 0m`/`E26 = 0m` manual y anti-tautología | D-H5.4, S4.3/CA-C5 |
| H5.5 | Tabla corpus-período (par R10+Administrativo por período) + gate R10 como malla mínima documentada | D-H5.5, S4.4 |
| H5.6 | Checklist de simetría con veredicto por familia (R1-L, R2-Especiales, R4-Especiales, Deb/Cred-otros, BCE) | D-H5.6, S4.4 |
| H5.7 | Suite trx + set-diff final; review independiente **APROBADO**; doctrina en `.opencode/project-context.md` (solo lo probado); cierre del plan con desviaciones fechadas | S4.5, CA-C6 |

**Orden sugerido:** H0 → H1 → H2 → H3 → H4 → H5. Regresión continua (suite + huella en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero). Excel cerrado durante las corridas.

---

## §5. GATES de cierre (medibles — todos o no cierra)

- G1: `dotnet build` → **0 warnings / 0 errores**; suite verde con set-diff exacto vs base H0 (rojo fuera de lista = STOP).
- G2: **2026091 ASE2 = 14.147.357.755** y **2026092 ASE4 = 7.675.362.502** (o valores T0 si §0 difiere), cada uno con Excel `(sin validar)` entregado + TXT de novedades.
- G3: ambos períodos producen TXT cuando corresponde; **ningún caso entrega nombre limpio con novedades**; ningún camino reporta `OK`/0/verde con novedades.
- G4: **0 regresiones** en 2026071/2026072/2026082/2026101 (R4 byte-idéntico donde la columna no existe; resto intacto).
- G5: candados H5.1/H5.2/H5.3/H5.4 verdes (incluidos injertos en rojo: encabezado nuevo no mapeado, letra fija, `P9 = 0m` manual, assert tautológico).
- G6: huella v4 intacta; review independiente **APROBADO**; UX del TXT ratificada por el Ingeniero.

---

## §6. Clarification Gate (requiere aprobación del Ingeniero; sin esto no se abre H0)

**Decisiones de negocio/UX (las toma el Ingeniero):**

1. **D-NOVEDAD-27:** ¿la validación cruzada 2.7 no declarada (`VALIDACION_TOTAL` FALSE en el workbook) degrada a novedad con Excel + TXT (recomendado §2-D-H3-i), o sigue siendo fallo duro sin entrega?
2. **D-TXT:** ¿un solo reporte siempre (` - Novedades de la corrida.txt`, retirando el nombre anterior) o compatibilidad (nombre ` - Divergencias declaradas.txt` cuando SOLO hay declaradas)? Recomendado: compatibilidad salvo que H0 demuestre cero dependencia del nombre.
3. **D-CODIGO:** ¿código de salida **7** para entrega-con-novedades (nombre propuesto `NovedadesEntrega`)? ¿El orquestador del lote lo trata como "revisar y seguir" o como fallo?
4. **D-DURO-ARCHIVO:** en fallo duro con `sin validar` ya escrito, ¿el archivo queda (evidencia, comportamiento actual) o se elimina (fallo limpio)? Recomendado: queda (cero cambio de comportamiento).
5. **D-UX-TXT:** ¿audiencia y profundidad de la guía "cómo revisar" (hoja/celda vs Administrativo/R10 por etiqueta)? El H4 trae muestra real para ratificar lenguaje.
6. **D-CORPUS:** ¿todo período nuevo entra con su par R10 + Administrativo como oráculos (disciplina H5.5)?

**Decisiones técnicas (las ratifica el orquestador/Ingeniero):**

7. **D-MODELO:** ¿`ReversionR4` se extiende aditivamente (`EspecialesReversion`, `TotalReversiones` = D observado) o cambia la semántica de `TotalReversiones` a D−P? Recomendado: aditivo (observación fiel intacta).
8. **D-ENCABEZADO:** ¿tabla única `EncabezadosEspeciales` compartida entre `ResolvedorConciliacionEmpresa`, `MapearR4` y `LeerR4` (recomendado, un solo punto de alias) o cada reader con la suya?

### Ratificación del Ingeniero (2026-10-28)

Las 8 decisiones se ratifican **con la recomendación senior del plan**. Vinculantes para H0..H5:

| # | Decisión | Ratificado |
|---|---|---|
| D-NOVEDAD-27 | La 2.7 no declarada (`VALIDACION_TOTAL` FALSE en el workbook) **degrada a novedad** con Excel + TXT | ✅ Novedad |
| D-TXT | Nombre del reporte: ` - Novedades de la corrida.txt` cuando hay novedades de datos; ` - Divergencias declaradas.txt` se conserva cuando SOLO hay declaradas | ✅ Compatibilidad |
| D-CODIGO | Código de salida **7** (`NovedadesEntrega`), distinto de 0 y de 6; el lote lo trata como "revisar y seguir" | ✅ 7 = revisar y seguir |
| D-DURO-ARCHIVO | En fallo duro, el `(sin validar)` ya escrito **queda** como evidencia (comportamiento actual) | ✅ Queda |
| D-UX-TXT | Guía "cómo revisar": hoja/celda contra Administrativo y R10 **por etiqueta**; muestra real se ratifica en H4 | ✅ Por etiqueta |
| D-CORPUS | Todo período nuevo entra con su par R10 + Administrativo como oráculos (disciplina H5.5) | ✅ Obligatorio |
| D-MODELO | `ReversionR4` se extiende **aditivamente**: `TotalReversiones` sigue siendo el D observado; se agrega el Especiales | ✅ Aditivo |
| D-ENCABEZADO | Tabla **única** de alias de encabezado compartida por `ResolvedorConciliacionEmpresa`, `MapearR4` y `LeerR4` | ✅ Única |

Sin 1-8 no se abre H0; sin H0 verde no hay H1; sin candados ni review no hay cierre.

---

## §7. Riesgos

| ID | Riesgo | Probabilidad / Impacto | Mitigación + gate que lo detecta |
|---|---|---|---|
| R-LETRA | Fijar la letra P (o P/Q por ASE) en vez del encabezado | Media / Crítico | D-H1 + H2.2 (columna movida) + candado H5.2; grep G-QUEMADO |
| R-UN-LADO | Corregir un solo lado leaf/agregado | Media / Alto | S2.1 (mismo cambio) + coherencia con Especiales ≠ 0; G2 |
| R-FALSO-VERDE | Novedad reportada como `OK`/verde/nombre limpio | Baja / Crítico | D-H3-iii + H3.4 (tests negativos) + CA-C3/G3 |
| R-REPORTE-JERGA | TXT ininteligible para el administrativo | Media / Alto | D-H4 (negocio, precedente Plan 42) + D-UX-TXT + CA-C4 |
| R-DOBLE-REPORTE | TXT de novedades vs TXT de declaradas contradictorios | Baja / Alto | D-H4 (un solo reporte con secciones) + H4.1 |
| R-ORACULO-FILA | Comparar por fila contra oráculos de geometría distinta | Alta si no se prohíbe / Alto | Prohibido normativo; gates por etiqueta; review lo exige |
| R-ORACULO-STALE | `<v>` de fórmula como verdad | Media / Alto | G-WORKBOOK; H0 por bytes |
| R-REGRESION | Mover períodos históricos que cierran | Media / Alto | G4 + H2.4 + set-diff (rojo nuevo = STOP) |
| R-REGRESION-EXIT | Algún test/herramienta asume exit ≤ 6 o nombre limpio | Media / Medio | T0.6 + H3.3 (preservar intención, no texto) |
| R-SUPUESTO-VIVO | Otro "P = 0" latente en otra familia | Media / Alto | H5.3 (barrido+clasificación) + H5.6 (checklist) |
| R-COLUMNA-FUTURA | Próximo mes trae otra columna intermitente ignorada | Alta / Alto | H5.1 (sonda: lo no mapeado es hallazgo, nunca silencio) |

---

## §8. Contradicciones entre el brief y el código real (resuelve H0 o escala)

| ID | Brief dice | Código real | Tratamiento en el plan |
|---|---|---|---|
| C1 | `ConciliacionEmpresaCierreBloqueT0Plan55Tests.cs:682` documenta "D − P, con P = 0" | El archivo con ese nombre **no existe**; el real es `ConciliacionEmpresaCierreBloqueT0Tests.cs` (sin `Plan55`), con el supuesto en `:682` (+ `:395`, `:821` con el mismo "P = 0") | H1.3/H0-T0.5 usan el nombre real; el fondo se confirma (causa §0.2-c en pie) |
| C2 | ValidadorBasico "línea ~886" y gate "solo en Q1 (línea 866)" | Exacto: mensaje en `:884-886`, early-return Q1 en `:866` con residuo `R-SIGMA-Q2` (`:853-865`) | Sin contradicción de fondo; el plan cita líneas exactas |
| C3 | "3 fixtures construyen `P9 = 0m` a mano" (lista cerrada) | Hay **dos más** no citados: `ValidadorBasicoTests.cs:90` y `ValidadorMultiAseTests.cs:174` (ambos `P9 = 0m`); además `E26 = 0m` convive en los mismos fixtures | T0.5 hace el inventario completo; H2.3 los cubre a todos |
| C4 | Corpus: "barrido de 30 archivos R4" … "solo 1 de 25 R4 tenía la columna" … "solo 2 tienen la columna" | Tres cifras distintas en el mismo brief (30 / 25 / 2) | T0.4 hace el barrido propio con conteo cerrado; el T0 manda |
| C5 | `ProcesadorPeriodoTests.WriterMultiAse_…_NoGeneraArchivo` y T19 "esperan ausencia de archivo" | Verificado: el test exige `!File.Exists(salida)` tras coherencia rota; `SalidaSinValidarO4Tests` y `AnclasDerivadasH2:462` YA contemplan el `sin validar` | H3.3 re-ancla esos tests preservando intención (no-certificar ≠ no-escribir); T0.6 los inventaría todos |
| C6 | "El gate R10 ya explota" el corpus por período | `ValidarDetRetriContraR10` hoy es fail-fast DURO pre-escritura (`:389` antes de `:402`): no "explota" nada, aborta sin Excel | D-H3 lo convierte en novedad post-escritura; H5.5 lo declara malla mínima |

---

## §9. No-objetivos (NO implementar)

1. No tocar la plantilla (cero `<f>`, cero `InsertRow`/`DeleteRow`, cero geometría; huella intacta).
2. No tocar cálculo, CONSOLIDADO, `DetRetri`/`DetValiRetri`, `VALIDACION_TOTAL` ni la cadena de fórmulas.
3. No extender la regla a ninguna otra hoja/columna fuera de R4-Especiales (H5.6 audita, no implementa).
4. No reabrir el predicado de canal, la identidad A/B, la delimitación por cierres (Plan 59), la reclasificación (Plan 60), ni Plan 58-P2.
5. No re-congelar letras, filas, ASE, períodos ni quincenas bajo ninguna forma (ni mapa, ni allow-list, ni "solo este ASE").
6. No incluir código productivo completo en este plan (solo firmas/pseudocódigo §2). No UI nueva salvo H4.3. No commits.

---

## §10. Trazabilidad con el brief (frente → dónde se cumple)

| Frente del brief | Dónde se cumple (sección + test) |
|---|---|
| F1: fix `MapearR4` + `LeerR4` + coherencia por encabezado | D-H1 + H1.1-H1.3 + CA-C1/G2 (casos 2026091/2026092 al centavo) |
| F1: asimetría empresa-vs-bloque cerrada | D-H1.3 (mismo cambio lógico) + Σ-empresas como malla en ambos períodos |
| F1: análisis de causa sistémica (a)-(e) | §0.2 + H5.1-H5.5 (cada causa tiene su candado) |
| F2.1: sonda de forma → hallazgo `ASE n · Reporte · Hoja · Encabezado` | D-H5.1 + H5.1 (reutiliza `EtiquetasNoReconocidas`) |
| F2.2: candado derivado de la plantilla (Total−Especiales) | D-H5.2 + H5.2 (lee la v4, no lista congelada) |
| F2.3: prohibir supuestos de ausencia + candado anti-literal | D-H5.3 + H5.3 + T0.5 |
| F2.4: doctrina de fixtures (no inventar fuente, no tautologías) | D-H5.4 + H2.3/H5.4 |
| F2.5: corpus por período con R10 + Administrativo | D-H5.5 + H5.5 |
| F2.6: auditoría de simetría (R1-L, R2-Especiales, R4) | D-H5.6 + H5.6 (checklist con veredicto) |
| F3: SIEMPRE Excel + TXT legible | D-H3 (orden + clasificación) + H3.2/H3.5 + G2/G3 |
| F3: clasificación novedad vs fallo duro | D-H3 (criterio general + tabla) + H3.2 |
| F3: modelo de resultado aditivo + reporte + nombres + superficies + códigos | D-H3/D-H4 + H3.1/H4.1-H4.3 + CA-C3/C4 |
| F3: prohibición de falso verde + riesgos de regresión | D-H3-iii + H3.3/H3.4 + T0.6 + §8-C5 |

---

## §11. Estado

**APROBADO 2026-10-28** por el Ingeniero. Clarification Gate §6 ratificado 8/8 con la recomendación senior (ver «Ratificación del Ingeniero»). **H0 habilitado**; H1..H5 quedan sujetas a la puerta H0→H1 (fixture verde + APROBADO) y a las puertas de cada fase.

---

## §12. Desviaciones ratificadas en H0 (2026-10-28)

H0/F0 cerrado: fixture `Remuneracion.IntegrationTests/EspecialesR4T0Plan61Tests.cs` (17 tests verdes) + evidencia `Docs/T0-Plan61-Especiales-R4.md`. **Todos los montos y celdas del §0 se re-midieron y coinciden al centavo** (2026092 ASE4: D23=−178.619.709,53 / P23=−49.900,44, R10 7.675.362.502; 2026091 ASE2: D11=−8.739.602,37 / P11=−26.500,00, R10 `DetRetri2026091!D10`=14.147.357.755). No hubo STOP de montos.

| ID | Hallazgo de H0 | Decisión ratificada (senior) |
|---|---|---|
| **D1** | No existe la hoja `CONSOLIDADO`; la real es **`CONSOLIDADO_TOTAL RECAUDO`** | Se corrige la cita en todo el plan; los gates y tests nombran la hoja real. Sin cambio de alcance |
| **D2** | No existe hoja `DetRetri`; son **`DetRetri{AAAAMMQ}`** (`DetRetri2026091`, `DetRetri2026092`) | Se corrige la cita; el lector de oráculo ya resuelve la hoja por período (`HojasPeriodo`). Sin cambio de alcance |
| **D3** | El árbol trae **25 rojos pre-existentes**; el trx del gate de entrada quedó obsoleto (3 commits del mismo día añadieron 39 tests y arreglaron 9) | Cada H mide **set-diff contra el árbol al abrirse**, no contra un trx viejo. Los 25 rojos se registran como **deuda declarada** (fuera del alcance del Plan 61, no se toleran rojos nuevos) |
| **D4** | Huella del frontmatter: `…a524d…` vs constante canónica `…a526d…` | Corregido el frontmatter (D4 aplicado). Gate = constante `HuellaSemanticaPlantilla`, nunca un literal del plan |
| **D5** | `ProcesadorPeriodoTests…_NoGeneraArchivo` y `VerificadorRecaudo` T19 **no pasan por `ProcesadorPeriodo`**: llaman a `OpenXmlPlantillaWriter.GenerarWorkbook` directamente, y los gates de coherencia viven **dentro del writer** (`OpenXmlPlantillaWriter.cs:240-246`), que aborta antes de crear archivo. Reordenar solo el procesador NO deja el `(sin validar)` | **Se adopta (b) con partición**: se **sacan del writer al procesador** los gates de **divergencia de datos** (`WorkbookLeafCoherence` leaf-vs-agregado) → degradan a novedad; el writer **conserva** los guardas **estructurales** (destino con fórmula, hoja/celda inexistente, plantilla alterada, límites de bloque), que siguen siendo fallo duro. H3.3 re-ancla los dos tests preservando la intención: *no certificar ≠ no escribir* (debe existir `(sin validar)` y **no** existir el nombre limpio). El criterio es el del §2-D-H3, no un caso especial |

**Puerta H0→H1: ABIERTA.** S1.1–S1.6 cerrados; G-H0-READONLY verificado (`git status` de producción vacío); build 0 warnings; §8-C1..C6 resueltas. H1 puede abrirse con D1..D5 aplicadas.

---

## §13. Cierre de H1 y enmienda propuesta (2026-10-28)

### 13.1 H1 ejecutado (verificado con salida real)

`MapearR4` + `LeerR4` + `WorkbookLeafCoherence` resuelven Especiales por el encabezado único. **Los dos números objetivo se reproducen al centavo**:

| Caso | Antes | Después | Esperado | Exit |
|---|---|---|---|---|
| 2026092 ASE 4 | 7.675.312.601,52 | **7.675.362.501,96** | 7.675.362.501,96 | ✅ número / ⚠️ bloqueado por 13.3-D |
| 2026091 ASE 2 | 14.147.331.255,39 | **14.147.357.755,39** | 14.147.357.755,39 | ✅ número / ⚠️ bloqueado por 13.2-C |

Suite: **0 rojos nuevos** (25 pre-existentes intactos, nombre por nombre). Build **0 warnings**. Cero cambios en la plantilla; `G-QUEMADO` limpio en el diff.

**Desviaciones de H1 que se ratifican (ya ejecutadas, inevitables):**

| ID | Desviación | Motivo | Veredicto |
|---|---|---|---|
| **D6** | `Core/Services/CalculoRemuneracion.cs:40`: `ReversionR4 = TotalReversiones − EspecialesReversion` (1 línea) | `ConsolidadoAse.ReversionR4` se llenaba con el Total bruto; con P≠0 saltaban `ValidadorBasico:835` (Q1) y `WorkbookLeafCoherence:200-203`. Es la misma forma de R2 (`TotalOportuno = GrandTotal − ServEspK`). La alternativa (sumar el especiales al consolidado) dejaba `TotalAse` mal | ✅ Ratificada. No es "tocar el cálculo": es que el consolidado lleve el **neto** que la plantilla publica |
| **D7** | Re-anclaje de 2 aserciones del fixture T0 (`EspecialesR4T0Plan61Tests`: `T0_4`, `T0_5`) | Fotografiaban el defecto (`var p9 = 0m;`, "sin análogo"/"patrón histórico"); mueren por construcción al corregirlo. Se conservó la intención: ahora afirman resolución por encabezado y dejan `WorkbookLeafCellMapQ2.cs` como deuda declarada para H5.3 | ✅ Ratificada. Un fixture de caracterización de un bug **debe** cambiar cuando el bug se corrige |

### 13.2 Bloqueo C — 2026091 ASE 3, Δ = 1 peso: **error de transcripción del oráculo**

Cadena probada: el R10 `DetRetri2026091!D11` es **texto** (`t="str"`, 116/116 celdas sin `<f>`), tecleado a mano, y su valor `…740` **no se reproduce con ninguna regla de redondeo** (ASE2 .39→755 y ASE5 .33→378 refutan CEILING/ROUNDUP). El administrativo calcula `D106 = 10.891.338.739,48` (Σ exacta de sus 5 bloques) → `ROUND` = **739**. Además: **el R10 es copia formateada del `DetRetri` del administrativo**, celda por celda — no es un oráculo independiente. Dueño: **insumo/oráculo (errata humana)**. La app y la plantilla están correctas.

**Decisión nueva propuesta — D8 (negocio):** *cuando el R10 y la cadena de fórmulas de la plantilla discrepan, manda la plantilla; la diferencia se reporta como novedad y NUNCA se "corrige" el cálculo ni se amplía la tolerancia.* Salida definitiva: la administradora corrige `D11` (740→739) y `D14` (623→622).

### 13.3 Bloqueo D — 2026092 ASE 2, `ERR-ESCRITURA`: **mapa de destinos congelado e incompleto en el writer**

Cadena probada: `WorkbookLeafCellMapDetalleR2R4.ColumnasR4` (`:145-160`) declara **13** pares congelados y **no incluye** `SERVICIO ESPECIALES` ni `Ajuste Recargo Aseo`. `EscribirDetalleR4` (`OpenXmlPlantillaWriter.cs:2182-2186`) escribe **por nombre contra esa tabla**: el valor no se desplaza, **se descarta**. El verificador sí deriva las columnas de la fila de encabezados de la hoja (`ColumnasDeValor:1414-1450`) y por eso lo caza. La plantilla v4 tiene destino de valor para ambos (P y Q) y fórmulas que los consumen (`CONSOLIDADO!AC67 = Q+R`, `AB69 = N69`). Dueño: **APP**.

Barrido: **6 de 30** fuentes R4 traen un encabezado fuera del mapa; **3 de los 10 R4 de septiembre bloquean** (2026091 ASE2, 2026092 ASE2, 2026092 ASE4) y otros 3 pierden centavos **en silencio** (≤ `ToleranciaDetalle` 0,5): Julio-1 ASE5 (+0,02), Octubre-1 ASE2 (+0,16), Septiembre-2 ASE5 (−0,09). **Es una clase, no un caso.**

Corrección al T0: el corpus real es **6/30 con encabezado y 6/30 con valor ≠ 0** (el T0 contó solo el alias `SERVICIO ESPECIALES` e ignoró `Ajuste Recargo Aseo`). Se corrige el conteo y el comentario heredado.

**Enmienda D9 — RATIFICADA por el Ingeniero (2026-10-28): el ajuste va por el MANEJO CONCEPTUAL de la plantilla, no por listas de columnas.**

Premisa: la plantilla no es una rejilla de letras; es una jerarquía declarada. La identidad de un valor es **(hoja · bloque de ASE · sección · subsección · fuente/canal · concepto · componente)** y el destino se resuelve cruzándola contra las **etiquetas de la propia plantilla** (doctrina vigente de Planes 55/56/59: derivación por rótulo y rol, cero letras congeladas). El mapa `ColumnasR4` con 13 entradas fijas contradice esa doctrina: sustituye la identidad por una posición.

**H1.6 (antes de cerrar H2)** — ejecutar **H1.6** antes de H2:
1. **Resolver conceptual en la escritura de detalle R4** (y R2 si comparte el mismo patrón congelado — se mide y se decide con evidencia en la misma tarea):
   - **Fila destino** = dentro del bloque del ASE, la fila cuya identidad de rótulos coincide (`A` cuando porta empresa · `B` = sección: `Total`/`Componente`/`Subs/Cont` · `C` = concepto: `Vlr Servicio`/`Vlr Intereses`/cierre/`Total`).
   - **Columna destino** = dentro de la fila de encabezados del bloque, la columna cuyo **encabezado** coincide con el componente de la fuente (normalizado: `Trim` + `OrdinalIgnoreCase` + tabla única de alias).
   - **Cero letras y cero conteos** en producción: `ColumnasR4` deja de fijar destinos; en todo caso sobrevive como **vocabulario de alias**, nunca como posición.
2. **Garantía de no-pérdida (invariante dura):** si una fila de la fuente trae un valor ≠ 0 en un componente que **no resuelve** a un destino, se lanza fallo duro nombrando `ASE · sección · concepto · encabezado · valor`. Prohibido el descarte silencioso y prohibido el 0 por defecto. Esto es lo que hoy deja perder −10,16 / +0,16 / +0,02 / −0,09.
3. **Re-anclar** `RevisionPestanasLenteR1T0Plan47Tests.cs:286-293`, que hoy **afirma la ausencia** de los encabezados y congela `Assert.Equal(13, …)`: se reemplaza por el invariante conceptual (todo encabezado de la fuente con valor tiene destino resuelto; todo encabezado de la plantilla es alcanzable).
4. **Extender H5.2** al camino de escritura: `{encabezados de la plantilla} ⊆ {destinos resolubles}` y `{encabezados de fuente con valor ≠ 0} ⊆ {destinos escritos}`, derivado de la v4 en disco (sin listas).
5. El guard `VerificarDetalleR2R4` **sigue siendo fallo duro** (D5 confirmado): detecta un artefacto que no puede construirse fielmente, no una divergencia de criterio.
6. **Criterio de aceptación de la clase, no del caso:** los 6 archivos del corpus con 14.ª/15.ª columna (Jul-1 ASE5, Oct-1 ASE2, Sep-1 ASE2, Sep-2 ASE2/ASE4/ASE5) deben escribir su componente en la columna que la plantilla declara para ese encabezado. Los tres que hoy pierden centavos en silencio **cambiarán** su salida: eso es la corrección, y se declara explícitamente.

**Tensión abierta para H3 (D10):** el writer **borra el parcial** al fallar, así que un fallo duro posterior a la escritura deja al usuario sin Excel, en tensión con D-DURO-ARCHIVO. Propuesta: conservar el `(sin validar)` solo en los fallos de **lectura-de-vuelta de valores** y seguir borrando en `ERR-PLANTILLA`/`ERR-INTEGRIDAD`.

**Estado:** H1 CERRADO en números. **D9 RATIFICADA** (ajuste conceptual, no lista de columnas) → **H1.6 EJECUTADO Y CERRADO** (ver §13.4). D8 (manda la plantilla cuando el R10 discrepa) y D10 (qué fallo duro conserva el `(sin validar)`) se ejecutan en H3. §13.2-C se resuelve en H2.1 (tercera regresión) + H3.2 (clase = novedad); §13.3-D se resolvió en H1.6.

---

## §13.4 Cierre de H1.6 (2026-10-28)

**Qué se hizo.** Se eliminó el mapa congelado de destinos en la escritura de detalle. Nuevo `Remuneracion.Infrastructure/Excel/ResolvedorDestinoDetalleR2R4.cs`: resuelve cada celda por **identidad conceptual** — bloque de ASE · sección (`B` = `Total`/`Componente`/`Subs/Cont`) · concepto (`C` = `Vlr Servicio`/`Vlr Intereses`/cierre) · componente (encabezado normalizado con tabla única de alias) — contra los rótulos de la plantilla. Cero letras de columna y cero conteos en producción. `WorkbookLeafCellMapDetalleR2R4` deja de fijar posiciones.

**Invariante nueva (dura):** un componente de la fuente con valor ≠ 0 que no resuelve a destino ⇒ fallo duro nombrando `ASE · sección · concepto · encabezado · valor`. Prohibido el descarte silencioso y el 0 por defecto. Se añadió test de injerto que lo demuestra.

**Evidencia medida (primera mano, 2026-10-28):**

| Verificación | Resultado |
|---|---|
| Build `Remuneracion.WinForms.slnx` | **0 Warning(s) / 0 Error(s)** |
| **2026092** CLI 5 ASE | **`RESULTADO OK codigo=0`**, Excel generado; `ASE 4: DetRetri = 7675362502` (origen `7675362501.96`) |
| Detalle R4 escrito por concepto (ASE 2) | `Reversion Pagos R4!Q117 = -10.16` y `R117 = 289.19` — antes `Q` quedaba en 0 (componente descartado) |
| **2026091** CLI 5 ASE | `ERR-VALIDACION` **con un único fallo**: Bloqueo C (ASE 3, R10 740 vs calculado 739). Sin fallos nuevos |
| Suite completa (`p61-h16-suite.trx`) | 1685 tests, **24 rojos**: **0 ALTAS**, **1 BAJA** (`ParidadCliTests` — rojo de harness que dejó de reproducirse; queda a confirmar en H3, no se da por arreglado) |
| `RevisionPestanasLenteR1T0Plan47Tests` | re-anclado: ya no afirma la ausencia de los encabezados ni congela el 13 |

**Estado de entrega real:** 2026092 ya se entrega certificado. 2026091 **sigue sin Excel** porque el gate R10 es hoy fail-fast pre-escritura: eso lo resuelve **H3** (política de entrega). Alternativa inmediata para 2026091: que la administradora corrija la errata del R10 (`D11` 740→739, `D14` 623→622).

---

## §13.5 Cierre de H2 (2026-10-28)

**Solo tests.** Producción intacta (los archivos de producción en `git status` son de H1/H1.6).

| Entregable | Detalle |
|---|---|
| `EspecialesR4H2Plan61Tests.cs` (5 tests) | Casos 2026091-ASE2 (neto −8.713.102,37 → 14.147.357.755), 2026092-ASE4 (neto −178.569.809,09 → 7.675.362.502), **caso del peso** 2026091-ASE3 (739,48 → 739 vs R10 740, documentado como errata), `sin columna → 0` y `columna movida de letra` (P→T, resultado idéntico) |
| H2.3 | 5 `P9 = 0m` y 4 `E26 = 0m` retirados o marcados (`[SINTÉTICO]` / `[AUSENCIA-MEDIDA]`); tautologías `DetalleR2R4Tests:134,161` reemplazadas por medidos (−16.536.713,09 y −12.054.255,65) |
| Injerto | Esperado alterado ⇒ rojo con mensaje por etiqueta; revertido ⇒ 5/5 verde |
| Build | **0 warnings / 0 errors** |
| Suite | **1.690 tests · 1.666 passed · 24 rojos · ALTAS 0 · BAJAS 0** (`p61-h2-suite-final.trx`) |
| Procedencia de la ausencia | Barrido del corpus: Julio-1 no declara `SERVICIO ESPECIALES` en ningún ASE; sus D coinciden con los fixtures |

**Supuestos residuales declarados (a H5, no a H2):** `P9=0` en 3 fixtures como ausencia medida de 2026071; composiciones R2 sintéticas marcadas; tautologías R2 de `DetalleR2R4Tests` (~133/~160); inventario completo de valores inventados = H5.3/H5.4.

**D10 — RATIFICADA (senior):** en fallo duro posterior a la escritura, el `(sin validar)` **se conserva** solo en los fallos de **lectura-de-vuelta de valores** (`VerificarDetalleR2R4`, `VerificarAsientoDetRetri`), y **se sigue borrando** en `ERR-PLANTILLA` / `ERR-INTEGRIDAD` (estructura rota no es evidencia útil).

---

## §14. Cierre de H3/H4 (2026-10-28)

**Entregado.** El pipeline quedó invertido: `preflight → lecturas → cálculo → escritura (sin validar) → gates de datos (novedad) → 2.7 → promoción solo con cero novedades`. Nuevas piezas: `EstadoEntrega`, `NovedadEntrega`, `ClasificadorNovedadesEntrega`, `InvarianteEntrega` (Core), `IReporteNovedades` + `ReporteNovedadesWriter` (Infrastructure), `CodigosSalida.NovedadesEntrega = 7`, superficie CLI y WinForms ámbar con botón de apertura.

**Verificado de primera mano (2026-10-28):**

| Corrida | Resultado |
|---|---|
| **2026091** | `RESULTADO CON NOVEDADES codigo=7` · Excel `(sin validar)` **620.235 B** + `Novedades de la corrida.txt` **3.433 B** · 1 novedad (ASE 3, 739,48→739 vs R10 740) |
| TXT | UTF-8 **sin BOM** verificado a nivel de bytes (3× `cál` correcto, 0 doble-codificado, em dash U+2014 presente) |
| **2026092** | `RESULTADO OK codigo=0`, nombre limpio, **sin** TXT |
| Anti-falso-verde | 3 negativos + injerto (`NovedadesEntrega = 0` ⇒ 2 rojos; revertido ⇒ verde) |
| D10 | lectura-de-vuelta conserva `(sin validar)`; `ERR-PLANTILLA`/`ERR-INTEGRIDAD` no |
| Build | **0 warnings / 0 errors** |
| Suite | **1.699 tests · 1.675 passed · 24 rojos · ALTAS 0 · BAJAS 0** (`p61-h3-suite.trx`) |

**Decisiones ratificadas en el cierre:**

| ID | Decisión |
|---|---|
| **D11** | **Modo 1-ASE (`--ase N`) queda FUERA del alcance** de H3: es un camino diagnóstico (Q2 no soportado) y sigue fail-fast. Se registra como **deuda declarada** con su propio plan corto. El camino de producción (5 ASE) sí entrega siempre |
| **D12** | **Declarada sin reconocimiento**: manda el flujo Plan 42 (excepción con clasificación + exit 1 para que opere el diálogo de reconocimiento), pero **después** de dejar Excel marcado + TXT. No se fuerza exit 7 ahí porque rompería el reconocimiento |
| **D13** | **UX del TXT**: se ratifica el texto entregado (lenguaje de negocio, "manda la plantilla", comparación por etiqueta). Ajustes de voz son cosméticos y no tocan lógica |
| **D14** | **Nombres**: ` - Novedades de la corrida.txt`, `CodigosSalida.NovedadesEntrega = 7`, `EstadoEntrega.EntregadaConNovedades`; `NovedadEntrega` usa `decimal?` para no inventar ceros y añade `Explicacion` |

**Incidente declarado:** durante H3 se corrompió y luego se truncó a 0 bytes un fixture *untracked* (`EspecialesR4H2Plan61Tests.cs`). Se recuperó verbatim del almacén de snapshots de OpenCode y se verificó con `git diff --no-index` que solo difiere en las ediciones intencionales. No hubo daño en repo ni en git.

**Pendiente:** H5 (sonda de encabezados, candado derivado de la plantilla extendido a la escritura, auditoría de supuestos, doctrina de fixtures), Plan 62 (triaje de los 24 rojos), D11 (modo 1-ASE) y la errata del R10 de 2026091 en manos de la administradora.

### D13-bis — Reporte escueto (ratificado por el Ingeniero, 2026-10-28)

El Ingeniero pidió un reporte **puntual, sin parafernalia**, y que el instructivo viva en el manual. Se retiraron del TXT: (1) el párrafo de presentación del estado; (2) el bloque **Datos de la corrida** (período, RunId, usuario, canal, fecha/hora, ruta del Excel); (3) la sección **«Qué significa "entregado con novedades"»**, incluida la regla «MANDA LA PLANTILLA»; (4) el párrafo **Regla general** de comparación por etiqueta; (5) el pie «no es un histórico».

**Contrato vigente del TXT:** veredicto `ENTREGADO CON NOVEDADES — NO CERTIFICADO` → `Detalle de las novedades (n)` con una ficha por novedad (ubicación, concepto, observado, esperado, Δ, qué pasó, dueño, qué revisar) → `Cómo revisar` (5 pasos) → sección de declaradas si aplica. Tamaño del caso 2026091: **3.433 → 1.564 B (−55 %)**.

La trazabilidad (RunId, usuario, canal, fecha) **no se pierde**: viaja en el log Serilog y en la salida de UI/CLI (`RESULTADO CON NOVEDADES ... runId=...`). Los parámetros de `IReporteNovedades.Emitir` se conservan por estabilidad del contrato y quedan documentados como «no se imprime».

---

## §14. Cierre de H1.6 — destino conceptual del detalle R2/R4 (2026-10-28)

**H1.6 EJECUTADO.** El destino de las columnas del desglose-detalle deja de ser una lista y pasa a ser una identidad resuelta contra la plantilla. **Los dos números del §13.3-D y los tres centavos se reproducen**, y el guard de valores sigue siendo fallo duro.

### 14.1 Qué cambió (traza a D9.1–D9.6)

| D9 | Implementación |
|---|---|
| **1 · fila destino** | Ya era conceptual desde el Plan 47: `EmparejarBloque` casa por `(firma A/B/C [R4] · A/B/C/D [R2] + canal efectivo)` contra los rótulos leídos de la hoja. H1.6 **no** la toca; el bloque se delimita con `DelimitadorBloquesHoja` (cabecera de bloque + gran total declarados) |
| **1 · columna destino** | Nuevo `ResolvedorDestinoDetalleR2R4` (Infrastructure): `ColumnasDeclaradas` lee la fila de encabezados del bloque desde la referencia A1 de cada celda; `IndicePorEtiqueta` normaliza (`Trim` + `OrdinalIgnoreCase`) y detecta identidad ambigua; `ColumnaDe` resuelve con la **tabla única de alias** `ResolvedorConciliacionEmpresa.EncabezadosEspeciales` (nadie duplica la lista) |
| **1 · cero posiciones** | `ColumnasR2` / `ColumnasR4` **retirados** de producción (eran 14/13 pares congelados). No sobrevive ninguna lista de posiciones; el "vocabulario" es el de alias ya existente |
| **2 · no-pérdida** | `EscribirComponentesDeFila`: componente con valor ≠ 0 sin columna declarada ⇒ `ERR-PLANTILLA` nombrando hoja · ASE · sección · subsección · concepto · encabezado · valor. Componente en 0/vacío sin destino = cero documentado (nada que perder). Dos encabezados de la fuente a la misma columna ⇒ `ERR-PLANTILLA` (nunca se pisa un valor) |
| **3 · re-anclaje** | `RevisionPestanasLenteR1T0Plan47Tests`: se eliminó `Assert.Equal(13, …)` + los dos `Assert.DoesNotContain` que bendecían la ausencia; ahora afirma alcanzabilidad (`todo encabezado declarado resuelve a SU columna`), destino de `SERVICIO ESPECIALES`→P, `Ajuste Recargo Aseo`→Q, `Otros`→R y el alias `Especiales`→P. Gemelos re-anclados: `DetalleR2R4EscrituraTests` (verifica componente a componente contra la hoja), `DetalleR2R4EscritoPlan59Tests.G6` (deriva el vocabulario de la v4) |
| **4 · H5.2 en escritura** | `VerificarDetalleR2R4` usa el MISMO resolvedor: compara Σ por COLUMNA destino (no por rótulo) y el bucle de destinos ausentes también resuelve por alias. Gate y writer ya no pueden discrepar |
| **5 · guard duro** | `VerificarDetalleR2R4` intacto: `ToleranciaDetalle` sigue en 0,5 y sigue abortando con `ERR-ESCRITURA` |
| **6 · la clase** | Barrido propio del corpus (30 fuentes R4, 30 R2): 6 R4 con 14.ª/15.ª columna y valor ≠ 0 (Jul-1 ASE5 +0,02 · Oct-1 ASE2 +0,16 · Sep-1 ASE2 −26.500,00 · Sep-2 ASE2 −10,16 · Sep-2 ASE4 −49.900,44 · Sep-2 ASE5 −0,09). **0 R2 con encabezado fuera de la plantilla** |

### 14.2 R2 — decisión con medición

R2 **sí compartía el patrón congelado** (mismo bucle `encabezado → letra`, mismo `?? 0m` por defecto), pero **no perdía nada hoy**: sus 14 pares coincidían exactamente con las 14 columnas que declara cada bloque de `Rem. Anticipos R2` y el corpus no trae encabezados fuera de esa lista. Veredicto: **corregido en la misma tarea** — la clase era la misma y el coste de dejarla era el próximo `Especiales`/`Deb-Cred-otros` que la plantilla añadiera. Medido: 30/30 fuentes R2 sin encabezado fuera de la plantilla; 0 cambios de valores en las corridas de regresión.

### 14.3 Corridas reales (salida medida, no citada)

| Caso | Resultado | Números |
|---|---|---|
| **2026092** (`--cinco-ase`) | **exit 0**, `RESULTADO OK` | `CONSOLIDADO_TOTAL RECAUDO!D107 = 7.675.362.501,96`; **`AC67 = 279,03`** (Q+R del cierre del ASE 2, `Ajuste Recargo Aseo` −10,16 + `Otros` 289,19); **`AB69 = −49.900,44`** (= `N69` = `'Reversion Pagos R4'!P` del cierre del ASE 4) |
| **2026091** (`--cinco-ase`) | exit 1 — **único** fallo: Bloqueo C (§13.2) | ASE 2 `DetRetri = 14.147.357.755` (calculado 14.147.357.755,39); ASE 3 Δ = 1 peso contra el R10 (errata del oráculo) y su consecuencia en la Σ. Ningún otro gate enrojece |
| **2026091 ASE 2** (`--ase 2`) | exit 0 | `Reversion Pagos R4!D9 = −8.739.602,37` y `P9 = −26.500,00` |
| **2026071** (Julio-1) | exit 0 | `Ajuste Recargo Aseo` escrito en Q (cierre ASE 5 = +0,02) ⇒ `CONSOLIDADO!AC70 = 0,02` (antes 0) |
| **2026101** (Octubre-1) | exit 0 | Q (cierre ASE 2 = +0,16) ⇒ `AC67 = 0,35` (antes 0,19) |
| **2026092 ASE 5** | exit 0 | Q (cierre = −0,09) ⇒ `AC70 = −0,09` (antes 0) |

**Cambio de salida declarado (D9.6):** los tres casos de centavos y los tres bloqueantes escriben ahora su componente en la columna que la plantilla declara. El cambio visible es en las columnas memo del CONSOLIDADO (`AB/AC`: «Conceptos especiales» y «Otros»); **ningún total (`D104:D108`, `D109`, `DetRetri`) se mueve por los centavos** — solo por el `SERVICIO ESPECIALES` del R4, que es el fix de H1 ya cerrado. No hay regresión en 2026071/2026072/2026082/2026101 (los cuatro siguen en exit 0 con los mismos totales).

### 14.4 Injerto de no-pérdida

Fixture `H16DetalleR4DestinoConceptualPlan61Tests.H162`: sobre una **copia** de `Docs/Prueba Septiembre-2/Insumos`, se injerta en la fuente R4 del ASE 4 el encabezado `Componente SINTETICO INJERTADO` con `12.345,67` en la fila `Vlr Servicio` y en el cierre. Resultado: **exit ≠ 0**, mensaje

> `ERR-PLANTILLA: la fuente de la hoja 'Reversion Pagos R4' · ASE 4 trae 12345.67 en el componente 'Componente SINTETICO INJERTADO' de la fila (sección '' · subsección '' · concepto 'Vlr Servicio') y la plantilla no declara ninguna columna de valor con ese encabezado en la fila de encabezados de su bloque. El monto NO se descarta ni se deja en cero…`

y **ningún `.xlsx` en la carpeta de salida** (no se entrega ni el limpio ni el `(sin validar)`).

### 14.5 Suite y build

- **Build**: 0 warnings / 0 errores.
- **Suite**: 1.685 tests (1.674 + 11 nuevos de H1.6) · **24 rojos** · 0 `Skip`.
- **Set-diff vs los 25 del §13 / D3**: **0 rojos nuevos**; **1 cambia a verde**: `ParidadCliTests.Paridad_5Ase_Q1_CliVsDirecto_ValoresYSnapshotsIguales`. Causa: su aserción mira el archivo `remuneracion_log_*.txt` del directorio de pruebas, que depende del orden con que los tests reconfiguran el `Log.Logger` global (`Program.ConfigurarSerilog`); en aislado pasa (5/5) y el fallo del trx base era «encontradas 0» líneas con el RunId, no un valor. **No es efecto de H1.6 y no es un rojo nuevo**; queda declarado como rojo histórico inestable (candidato a aislar el sink en un test propio).

### 14.6 Estado

**H1.6 CERRADO.** §13.3-D resuelto. Queda para H2.1 la tercera regresión de §13.2-C (errata del oráculo) y para H3 las decisiones D8 y D10. Sin cambios en la plantilla (huella v4 intacta), sin `<f>` nuevas, sin tocar cálculo/CONSOLIDADO/DetRetri/VALIDACION_TOTAL ni la política de entrega.

