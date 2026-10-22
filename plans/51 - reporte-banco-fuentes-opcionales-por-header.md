---
titulo: Plan 51 - Reporte banco: fuentes opcionales por header
estado: CERRADO 2026-10-22 (A1+A2+A3+A4 implementados; APROBADO CON OBSERVACIONES; W-1 corregido; S-1/S-2 declaradas)
fecha: 2026-10-22
numero: 51
planes_ocupados: 01..50 ocupados; este es el 51
padre: T0-Plan50 §4.4 (bloqueo #1 — dueño HU-09) + plans/50 (doctrina opcionalidad + A3-etiquetas) + plans/09 (contrato original T0-0.3/T0-0.7)
madre_doctrina: project-context.md (Modelo de la plantilla estándar; Plan 50: fuente ausente = 0 + huella; Plan 34: sinonimia; Plan 48: v2 0F192AD2…; Plan 28: values-only)
caso_rojo: corrida 2026101 (octubre-1, Q1) aborta en ASE 2 (LIME) — `CalculoInvalidoException — ASE 2: la columna fuente 2 del Resumen debería ser 'NUEVO ESQUEMA' pero dice 'OCCIDENTE'.` (reader `ExcelDataReaderWorkbookLeafInputReader.cs:257-262`, mapa posicional `WorkbookLeafCellMapReporteBanco.ColumnasFuentePorAse`)
gate_entrada:
  plantilla_v2: NO SE TOCA (SHA `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` intacto antes/después)
  plantilla_v1: NO SE TOCA (histórica; SHA `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` intacto antes/después)
  alcance_aprobado_pendiente: A1 (columnas por header) + A2 (huella) + A3 (desconocida ≠ ausencia) + A4 (writer/plantilla intactos); prohibido ampliar
  suite_base: p50-h2-suite.trx (1141 Total / 1074 Passed / 67 Failed / 0 Skipped)
  sin_commits: NUNCA `git commit`/`push`
---

# Plan 51 — Reporte banco: fuentes opcionales por header

> SDD completo PROPOSE → DESIGN → SPEC → TASKS. Fuente única de verdad para la implementación una vez aprobado ESTE documento. No se implementa sin aprobación del Ingeniero (punto de bloqueo: sin su OK no se abre H1).
> Lectura obligatoria previa: `.opencode/project-context.md` (Modelo de la plantilla estándar + doctrina Plan 50) + `plans/50` (precedente directo: O-C, W-A, A3-etiquetas, residuos) + `plans/09` (contrato original del banco, T0-0.3/T0-0.7) + `Docs/T0-Plan50-FuentesRecaudoOpcionales.md` §4.4 (e2e 2026101: el rojo banco ya probado).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Contexto adoptado (no re-investigar; cada hecho cita su fuente)

### 0.1 El rojo que motiva el plan (probado en el T0 del Plan 50, §4.4)

Corrida `2026101` (octubre-1, Q1) con A1 del Plan 50 ya verde aborta en ASE 2 (LIME):

```text
[ERR-VALIDACION] ASE 2: la columna fuente 2 del Resumen debería ser 'NUEVO ESQUEMA' pero dice
'OCCIDENTE'. Mapa T0-0.3 no coincide con la fuente.
```

Cadena causal (código vigente, verificado en esta planificación):

| Punto | Archivo:línea | Hecho |
|---|---|---|
| Mapa | `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapReporteBanco.cs:48-56` (`ColumnasFuentePorAse`) | ASE2 exige `ENEL=1, NUEVO ESQUEMA=2, OCCIDENTE=3` (**posicional**, congelado por T0-0.3 del Plan 09) |
| Resumen | `Docs/Prueba Octubre-1/Insumos/2-Lime/ReportePagosxBanco_...20261016174332439.xlsx`, `Sheet1`, R391-R396 | R391 `A=Resumen Recaudo Aplicado Por Servicio`; R392 `B=ENEL \| C=OCCIDENTE \| D=Total`; R396 `A=Total \| B=10.309.260.573 \| C=971.950.730 \| D=11.281.211.303` |
| Throw | `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs:255-262` | Header 2 es `OCCIDENTE`, no `NUEVO ESQUEMA` → `CalculoInvalidoException` que nombra ASE + columna |
| Writer | `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs:1329-1359` (`EscribirCeldasBanco`) | Itera **solo** `bloque.Empresas`; una fuente ausente no se escribe y queda el 0 de la plantilla (ya verificado: A4 no exige cambios salvo que el T0 demuestre lo contrario) |
| Gate | `Remuneracion.Core/Services/ValidadorBasico.cs:737-767` (`ValidarReporteBanco`, gate i) | Σ conceptos vs `TotalFuente` **por empresa presente**; el gate no se altera en este plan |
| Modelo | `Remuneracion.Core/Models/ReporteBancoInputs.cs` | `ReporteBancoEmpresaInputs` (4 conceptos + `TotalFuente` + `Total` calculado), `ReporteBancoAseInputs`, `ReporteBancoInputs` |

**Veredicto adoptado del T0-Plan50 §4.4:** es **la misma familia del Plan 50** — `NUEVO ESQUEMA` es una fuente de recaudo **opcional**; en octubre-1 no vino en el banco de LIME. El defecto es el mapa **posicional** T0-0.3, no los datos.

### 0.2 Alcance APROBADO para diseño (del encargo; prohibido ampliar)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **A1** | Columnas del resumen **por header reconocido**, no por índice congelado | Se recorre la fila de encabezados (tras `Resumen Recaudo Aplicado Por Servicio`) y se mapean solo los headers reconocidos (catálogo de fuentes + sinonimia vigente `NUEVO ESQUEMA`/`RECIPROCIDAD`). `Total` se ignora (no es fuente). Fuente del catálogo que **no viene**: no se exige, no se leen sus conceptos, el writer no la escribe, sus celdas quedan en 0 (la plantilla ya está en 0) |
| **A2** | Huella de ausencia del banco (misma doctrina Plan 50: campo aditivo auditable + 1 `Warning` por ASE+fuente, sin ruido) | Decisión §2.2: reutilizar `FuentesAusentes` con prefijo de reporte (`BANCO:`) vs campo propio — recomendado y justificado, a ratificar en §7 |
| **A3** | Etiqueta no reconocida ≠ ausencia | Header no-fuente-no-`Total` + falta una fuente del catálogo del ASE → **fail-fast** (posible cambio de etiqueta), nunca 0. Solo-subconjunto-válido → ausencia legítima (0 + huella) |
| **A4** | Writer y plantilla intactos | `EscribirCeldasBanco` ya escribe solo las empresas presentes (verificado §0.1); la plantilla v2 no se toca (SHAs arriba) |

### 0.3 Restricciones duras (vinculantes)

- **Plantillas intocables**: v2 `0F192AD2…` + v1 `3086AC31…`. Verificar antes/después.
- **Values-only**; cero fórmulas tocadas.
- **Tests no se debilitan**: 0 `Skip`, 0 `Assert.True(true)`, 0 borrados.
- **Build 0 warnings**, CRLF sin mezclar, sin `git commit`/`push`.
- **No tocar** el Plan 50 ni sus deudas (`R-MAPA-V2-R2-ASE4`, `R-PARCIAL-OCT-ASE3-R1-OCC`, `R-SIGMA-Q2`); no tocar el mapa por empresa de R1/R2/R4; no tocar el mapa posicional de otras fuentes sin T0.
- **Mapeo por ETIQUETA/HEADER**, nunca por posición/ocurrencia (doctrina transversal + Plan 09 D1).

### 0.4 Lo que el T0 (H0, read-only bloqueante) debe congelar

1. La fila de headers y las fuentes presentes/ausentes por **ASE × período** (julio-1, julio-2, agosto-2, octubre-1) en `ReportePagosxBanco`.
2. Que `Total` se distingue de una fuente (exclusión explícita, nunca vocabulario).
3. Que los conceptos (1/2/3/7) y la fila `Total` se siguen leyendo igual (etiquetas/prefijos intactos).
4. Que el `TotalFuente` por empresa presente sigue cuadrando (Σ conceptos) y que el gate (i) no se altera.
5. La decisión ASE2 (2 fuentes) vs ASE4/ASE5 (3) contra lo declarado por `EditablesPorAse`.
6. Cualquier header no reconocido en las muestras (si aparece, STOP y reporte).

---

## 1. PROPOSE

### 1.1 Intent

Que la lectura del `ReportePagosxBanco` deje de abortar cuando una **fuente de recaudo opcional** no viene en el período: el reader resuelve las columnas **por header reconocido**, la fuente ausente se **omite con huella auditable** (no es excepción, no es fila-cero inventada), y una **etiqueta desconocida que coexiste con una ausencia** sigue abortando en duro — mientras la estructura rota (etiqueta `Resumen` ausente, fila de headers ausente, cero fuentes reconocidas) sigue en fail-fast. Writer, plantilla, conceptos, fila `Total` y gate Σ: intactos.

### 1.2 In Scope

- H0: T0 read-only bloqueante (§0.4) + fixture que lo congela.
- H1: header-scan en `LeerReporteBanco` (retira el consumo de `ColumnasFuentePorAse`); huella A2 (modelo aditivo + propagación + 1 warning por ASE+fuente); verificación con evidencia de que el writer no cambia (A4); repro 2026101-ASE2 sin el throw posicional.
- H2: juicio A3-banco en `ValidadorBasico` + simulación (etiqueta mutada → fail; ausencia coherente → verde); regresión e2e + suite trx + set-diff vs base `p50-h2-suite.trx` (1141/1074/67/0) con lista cerrada.
- H3: doctrina en `project-context.md` (solo lo probado) + nota en `plans/09` (T0-0.3 superseded, sin reescribir historia).

### 1.3 Non-goals / out (vinculante)

- No se corrige ningún residuo del Plan 50 (`R-MAPA-V2-R2-ASE4`, `R-PARCIAL-OCT-ASE3-R1-OCC`, `R-SIGMA-Q2`) ni los 20 rojos de la v2 `d55e621`; no se promete el cierre e2e de 2026101 (§2.7: otros bloqueos quedan).
- No se tocan plantillas (v1/v2), ni fórmulas, ni sellados de caché, ni `DivergenciasDeclaradasCatalogo`, ni gates de otros planes.
- No se tocan los mapas por empresa de R1/R2/R4 ni el mapa posicional de ninguna otra fuente.
- No se reescribe `EscribirCeldasBanco` salvo que el T0 demuestre una divergencia (A4: verificación, no cambio).
- No se generaliza "ausencia = omitir" a R1/R2/R4 (allí rige Plan 50 W-A: 0 explícito porque el writer exige valor por celda).
- Commits, UI, CLI, docs de usuario.

### 1.4 Riesgos con mitigación

| ID | Riesgo | Mitigación |
|---|---|---|
| R-FALSO-0 | La fuente **viene** pero con etiqueta nueva y el header-scan la omite como "ausente" (falso verde) | A3 §2.3: desconocida + ausente en el mismo resumen → fail-fast; tercera forma §2.3.4 |
| R-TOTAL | `Total` se trata como fuente (se emite empresa `Total` o dispara A3) | Exclusión explícita por match normalizado `== "total"` **antes** del vocabulario + test dedicado (§2.3.1) |
| R-VACIO | Resumen sin NINGUNA fuente reconocida produce hoja en ceros silenciosa | Regla: cero reconocidas → fail-fast (estructura rota, no ausencia) (§2.3.4) |
| R-MAPA-DESYNC | Header reconocido pero sin celdas destino en `EditablesPorAse[ase]` (mapa↔plantilla desincronizados) | Se conserva el throw vigente `:1344-1348` (writer) / fail-fast en reader: bug real, nunca skip silencioso (§2.5) |
| R-FIXTURE | Re-congelar tests para "volver a verde" en vez de documentar el nuevo comportamiento | Re-congelamiento solo con evidencia citada (fuente sin la fuente + omisión esperada); lista cerrada; 0 `Skip`/borrados |
| R-RUIDO | La huella inunda el log y oculta errores reales | 1 `Warning` por (ASE, fuente) + campo agregado; sin volcado por celda (doctrina Plan 50) |
| R-Q1 | El cambio altera Q1 más allá del banco-no-aborta | G-Q1 (§3-S0): Q1 solo cambia en que la ausencia banco deja de abortar; set-diff con lista cerrada |
| R-CANCEL | El gate Σ por empresa (Total vs TotalFuente, ambos de la fuente) cancela un error sistemático | No se introduce operando nuevo: el gate sigue intacto y solo cubre empresas presentes; la ausencia no entra al gate (se omite, no se suma 0=0) |
| R-ALCANCE | El plan se expande al mapa R1/R2/R4, al A2 del Plan 50 o a Q2 | §1.3 vinculante + grep de cierre (ningún archivo fuera de la lista §2.6) |

---

## 2. DESIGN

### 2.1 D-1 — Fuente de verdad de los headers (decisión con trade-offs)

**Hecho de partida:** el vocabulario de fuentes del negocio ya existe y ya se usa para A3 en R1/R2/R4 — `EmpresaFacturacion.Catalogo` (`LabelTemplate`: `RECIPROCIDAD/ENEL/ENERBIT/OCCIDENTE/CIUDAD LIMPIA-ACUEDUCTO`) + `ObtenerLabelFuente(empresaId, aseId)` (rótulo vigente: empresa 1 en ASE 2/4 → `NUEVO ESQUEMA`; `ExcelDataReaderWorkbookLeafInputReader.cs:1664-1667`) + matcher puro `SinonimosEmpresaR1.SonMismaEmpresa` (clase `RECIPROCIDAD ↔ NUEVO ESQUEMA` con cita; catálogo ABIERTO: lo no tabulado solo se iguala a sí mismo). El Plan 50 §2.5 ya construye el vocabulario así (`:143-147`) y el T0-Plan50 lo probó.

| Candidata | Veredicto | Motivo |
|---|---|---|
| **D1-A: vocabulario global existente + esperado-por-ASE = claves de `EditablesPorAse[ase]`** | **Elegida (a ratificar §7.1)** | Cero catálogos nuevos (OCP/D-H): el reconocimiento usa el vocabulario citado del negocio (catálogo + rótulo vigente + sinonimia); lo **esperado por ASE** lo declara **la plantilla** (`EditablesPorAse`: ASE1/3 → 2 fuentes, ASE2/4/5 → 3), no una lista de columnas. `Total` se excluye explícito antes del vocabulario. La sinonimia ya probada absorbe `NUEVO ESQUEMA`/`RECIPROCIDAD` sin ramas |
| D1-B: lista propia del banco (nuevo catálogo `FuentesBancoPorAse`) | **Descartada** | Tercer catálogo que se pudre con cada forma nueva; duplica `EditablesPorAse` (que ya declara por ASE qué fuentes tienen celdas) y el vocabulario del Plan 50. D-H lo prohíbe en espíritu |
| D1-C: solo `EmpresaFacturacion.Catalogo`, sin rótulo vigente ni sinonimia | **Descartada** | `LabelTemplate` de la empresa 1 es `RECIPROCIDAD`, pero el banco trae `NUEVO ESQUEMA`: sin el rótulo vigente + sinonimia, la fuente real se volvería "desconocida" y A3 dispararía en falso (lección Planes 33/34) |

**Reglas normativas D-1:**

1. **Reconocimiento (global, sin ASE):** cada header no vacío y ≠ `Total` se resuelve contra el vocabulario global (los 5 `LabelTemplate` + los rótulos vigentes por ASE + sinonimia vía `SonMismaEmpresa`). El match usa **ese matcher** (mayúsculas sin diacríticos); la `NormalizarEtiqueta` local del banco (minúsculas, conserva `-`; `:374-386`) **no** se usa para empresas — dos normalizaciones distintas para el mismo juicio serían un bug latente. Para conceptos (filas) y `Total` (fila) se conserva `NormalizarEtiqueta` + prefijos `Conceptos` (intacto, §2.4).
2. **Esperado por ASE (plantilla manda):** `EditablesPorAse[ase].Keys` es el conjunto esperado. `ColumnasFuentePorAse` deja de consumirse (destino §2.5); sus índices físicos no se conservan ni como fallback (un fallback posicional reintroduciría el defecto).
3. **Ausente = esperada no observada** → se omite (no se leen sus conceptos, el writer no la escribe, celdas en 0 de plantilla) + huella A2.
4. **Desconocido = observado no reconocido y ≠ `Total`** → evidencia para A3 (§2.3), nunca invención: no se crea empresa, no se escribe nada por él.
5. **Cero ramas por período/empresa/fuente:** la ausencia se predica de la geometría observada (headers del resumen), nunca de un literal (`aseId == 2`, `"NUEVO ESQUEMA"`, `2026101`) en producción.

### 2.2 D-2 — Huella de ausencia (decisión con trade-offs)

| Candidata | Veredicto | Motivo |
|---|---|---|
| **D2-A: campo propio `FuentesAusentesBanco` en `ReporteBancoAseInputs` (`BANCO:{label}`) → propagación aditiva a `ResultadoProcesoPeriodo.FuentesAusentes` como `ASE {n} · BANCO:{label}` + 1 `Warning` por (ASE, fuente)** | **Elegida (a ratificar §7.2)** | Mismo canal durable ya auditado (log + resultado, formato `ASE n · X` intacto; UI/log no cambian de forma). Grano correcto: **1 bloque por ASE** (el banco no es por empresa de facturación, a diferencia de `ConciliacionEmpresaInputs` que replica ×5). `ValidarFuentesRecaudoNoReconocidas` (Plan 50) lee solo `leaf.Conciliacion`: el prefijo `BANCO:` no contamina su partición por reporte R1/R2/R4 |
| D2-B: reutilizar `ConciliacionEmpresaInputs.FuentesAusentes` | **Descartada** | Grano y familia equivocados: habría que replicar la huella banco en 5 entradas por empresa o forzarla en una; el validador del Plan 50 la leería como reporte `BANCO` dentro de una estructura R1/R2/R4 (confusión de dominio) |
| D2-C: campo propio nuevo en `ResultadoProcesoPeriodo` (p. ej. `FuentesAusentesBanco`) | **Descartada** | Duplica superficie durable (resultado + log + UI/CLI) para el mismo concepto ya modelado ("fuente ausente"); el formato `ASE n · X` lo absorbe sin cambios. Solo se reconsidera si el T0 demuestra colisión real de prefijos |

**Formato:** entradas `BANCO:{label}` con el label **observado del vocabulario** (p. ej. `BANCO:NUEVO ESQUEMA`). `ReporteDe("BANCO:X")` → `BANCO` por el split existente en `:` (verificar en H1 que el helper lo particiona igual; si no, ajuste mínimo citado).

### 2.3 D-3 — A3 banco: desconocida + ausente → fail-fast (sin falsos positivos)

**Regla (gemela del Plan 50 §2.5, grano reporte `BANCO`):**

1. **`Total` nunca es fuente ni desconocida:** la celda de header cuyo normalizado es `total` se excluye **antes** del vocabulario (test dedicado: el resumen octubre-1 ASE2 trae `D=Total` y no debe generar ni empresa `Total` ni evidencia A3).
2. **Celdas de header vacías se ignoran** (no son evidencia).
3. **Desconocida sin ausencia en el mismo resumen = solo evidencia** (gemelo `EAB` del Plan 50: etiqueta real sin destino no convierte en rojo lo sano).
4. **Desconocida + ausente (esperada no observada) en el mismo resumen = fail-fast** que nombra ASE + `BANCO` + etiqueta(s) desconocida(s) + fuente(s) ausente(s) (formato §2.6). Es la firma de "fuente presente no reconocida" (cambio de etiqueta).
5. **Cero fuentes reconocidas = estructura rota → fail-fast** (nunca hoja en ceros silenciosa). Una ausencia legítima presupone al menos una fuente reconocida leída con su `Total`.
6. **Sin umbral numérico** (lección Plan 50 §4.3/T0 §3: la aritmética Σ-presentes-vs-total no es viable como control; aquí además el gate Σ por empresa presente ya existe intacto y la ausente no entra al gate porque se omite).

**Evidencia (reader, sin juicio):** `EtiquetasBancoNoReconocidas: IReadOnlyList<string>` en `ReporteBancoAseInputs` (aditivo, formato `BANCO:{label}`), construida con las filas ya leídas (cero workbooks extra; prohibido `<v>` — no aplica aquí: todo es texto de header de fuente).

**Juicio (Core):** método nuevo en `ValidadorBasico` junto a `ValidarFuentesRecaudoNoReconocidas` (SRP: no se modifica el juicio R1/R2/R4 del Plan 50), invocado en el mismo punto; `ClasificarDivergenciasDeclaradas` lo hereda (última puerta, Plan 42 D-F: A3-banco disparado = NO declarada, falla en duro).

### 2.4 D-4 — Conceptos, fila Total y gate: intactos (confirmación)

- **Conceptos:** la lectura es por **índice de concepto** (`Conceptos` 0..3: `1-aplicadosafacturacion`, `2-saldosafavor`, `3-aplicadosafinanciaciones`, `7-aplicados`) sobre **filas** del resumen (`:300-313`). El cambio solo remapea **qué columna** se lee por empresa reconocida; el contrato fila→concepto no se toca. La fila 7 ausente en Q1 (= 0 demostrable, `ReporteBancoInputs.cs:34-38`) sigue igual.
- **Fila `Total`:** se sigue detectando por etiqueta normalizada `== "total"` (`:290-298`) y aporta `TotalFuente` por empresa **presente**. La ausente no tiene `TotalFuente` (no se emite fila: no hay `0` que confundir con un total real).
- **Gate (i) `ValidarContraFuentesBanco` / `ValidarReporteBanco`:** intacto en código y umbral (±0.5). Solo ve empresas presentes; la coherencia ausente-no-entra-al-gate es por construcción (omisión, no 0=0).
- **C59 y `Quincena`:** intactos (dominio, nunca fuente; Requirement 4 del Plan 09).

### 2.5 D-5 — Writer intacto + destino de `ColumnasFuentePorAse` (decisión)

**Verificación A4 (el T0 la confirma, H1 la demuestra con test):** `EscribirCeldasBanco` (`:1329-1359`) itera solo `bloque.Empresas` y consulta `editables.TryGetValue` **solo para las presentes**; con omisión de ausentes no hay camino que lance por ausencia. C59 se escribe siempre. Conclusión: **cero código en el writer**.

**Diferencia explícita vs Plan 50 W-A (justificada):** allí el 0 explícito era obligatorio porque `EscribirCeldasConciliacion` **exige** valor por celda del mapa (`:1287-1316`); aquí el writer itera presentes, luego la **omisión** es más fiel (no inventa filas ni `TotalFuente=0`) y deja el 0 de la plantilla. La huella vive en el campo A2, no en filas-cero.

| Candidata | Veredicto | Motivo |
|---|---|---|
| **Retirar el consumo de `ColumnasFuentePorAse` en el reader; borrar el mapa tras paridad T0** | **Elegida (a ratificar §7.5)** | Mapa muerto con índices físicos se pudre y tienta como fallback posicional (la clase de defecto que este plan cierra). El T0 demuestra paridad esperado(`EditablesPorAse`)↔observado antes de borrar |
| Conservarlo como `[Obsolete]` / fallback | **Descartada como principal** | Un fallback posicional reintroduce el defecto por la puerta trasera; `[Obsolete]` sin consumidor es ruido. Solo se acepta si H1 encuentra un consumidor legítimo (citarlo) |

**Regla conservada:** empresa reconocida **sin** celdas en `EditablesPorAse[ase]` → fail-fast (se conserva el throw `:1344-1348` o su gemelo en el reader): desync mapa↔plantilla = bug real, nunca skip.

### 2.6 Superficie de cambio (contract-first; lista cerrada)

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` (`LeerReporteBanco`, `:231-339`) | Modificar: header-scan D-1 + omisión + huella/evidencia | Núcleo A1/A2/A3-evidencia |
| `Remuneracion.Core/Models/ReporteBancoInputs.cs` | Modificar (aditivo): `FuentesAusentesBanco` + `EtiquetasBancoNoReconocidas` en `ReporteBancoAseInputs` | Huella y evidencia A2/A3 |
| `Remuneracion.Core/Models/ResultadoProcesoPeriodo.cs` | Modificar (aditivo, propagación): `FuentesAusentes`/`EtiquetasNoReconocidas` absorben `ASE n · BANCO:…` | Canal durable existente |
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` (`:560-569`) | Modificar (aditivo): agregar huella/evidencia banco al agregar las existentes | Propagación |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Agregar: juicio A3-banco junto a `ValidarFuentesRecaudoNoReconocidas` | A3-juicio (SRP, Plan 50 intacto) |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapReporteBanco.cs` | Modificar: retirar `ColumnasFuentePorAse` (destino §2.5) | Fin del mapa posicional |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs` | Intacto (verificado; gate i sin cambios) | — |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Intacto salvo prueba en contra del T0 (A4) | — |
| Tests nuevos | `ReporteBancoHeadersT0Plan51Tests` (H0) + `ReporteBancoFuentesOpcionalesPlan51Tests` (H1) + `ControlBancoNoReconocidoPlan51Tests` (H2) | Rojo→verde honesto |

**No tocar (grep de cierre):** `ColumnasFuentePorAse` fuera de §2.5; mapas R1/R2/R4 por empresa; `EscribirCeldasBanco` (salvo T0 en contra); plantillas; `DivergenciasDeclaradasCatalogo`; Plan 50 y sus residuos.

### 2.7 e2e 2026101: qué otros bloqueos quedan (para no prometer un cierre ajeno)

Tras este plan, la lectura banco de ASE2-2026101 deja de abortar. Los bloqueos restantes del período objetivo tienen dueño propio (T0-Plan50 §4.4 + residuos Plan 50) y **no** los cierra este plan:

| # | Bloqueo | Dueño |
|---|---|---|
| 1 | R1 ASE3 OCCIDENTE parcial (`TotalMedio` sin fila) — `R-PARCIAL-OCT-ASE3-R1-OCC` | T0 de negocio |
| 2 | Gate Σ: R4 ASE1 `ENEL` sin celda destino + R2 `EAB` — `R-SIGMA-Q2` (acotado Q1) | Mapa HU-08 (cobertura) |
| 3 | `ERR-SIN-SLOT` del contrato de sección R1 en fuentes Q1 | Planes 41/46 (Q1) |
| 4 | A2 Plan 50 (recorte Q2) — `R-MAPA-V2-R2-ASE4` | T0 mapa R2 (1,4) v2 |
| 5 | 20 rojos pre-existentes del reemplazo v2 `d55e621` | T0 propio de re-congelamiento |

### 2.8 Architecture Validation Certificate (diseño)

| Principio | Veredicto |
|---|---|
| **SRP** | Reader = evidencia (reconoce, omite, reporta huella/evidencia) / Validador = juicio (A3-banco + gates existentes) / Mapas = datos (`EditablesPorAse` declara lo esperado; el índice posicional sale) / Writer no juzga (itera presentes) |
| **OCP** | Fuente nueva o ausente nueva = cero código (header-scan + vocabulario abierto absorben por geometría observada). Solo un caso parcial/celda-faltante probado por T0 añadiría comportamiento, como dato citado |
| **DIP** | Cambios vía interfaces existentes (`IWorkbookLeafInputReader`, `IValidador`, `IPlantillaWriter`); `ProcesadorPeriodo` sigue dependiendo de abstracciones; campos aditivos, contratos intactos |
| **Best practices** | Contract-first (H0 congela headers/ausencias/reglas antes del juicio); fail-fast accionable (ASE + BANCO + fuente/etiqueta); tolerancia ±0.5 intacta; values-only; D-H (cero literales de período/empresa/fuente en producción); CRLF; sin commits |
| **Performance** | O(columnas del header + filas ya leídas); sin workbooks extra; el control A3 es O(ASE × header) en memoria; regens a temp fresca |

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda T)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa con trx + **set-diff vs línea base `p50-h2-suite.trx` (1141/1074/67/0)**: altas/bajas una por una con causa citada. **Rojo nuevo fuera de lista = STOP con causa.** 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados.
- G-SHA: v1 `3086AC31…C5214` + v2 `0F192AD2…2DAB` intactos antes/después.
- G-PROC: cero commits; CRLF sin mezclar en lo tocado.
- G-D-H: cero literales/ramas por período, empresa o fuente en producción (grep de cierre; los labels son datos del catálogo/vocabulario existente, nunca ramas nuevas).
- G-WORKBOOK: ningún gate abre workbooks extra para A3; prohibido `<v>` stale (todo es texto de header de fuente + literales ya leídos).
- G-FORMULA: `<f>` byte-idéntica en todo el diff (cero `new CellFormula`, cero `InsertRow`/`DeleteRow`).
- G-TRAZA: cada error nuevo cumple el formato §2.6/§3-S1 (ASE + `BANCO` + fuente/etiqueta).
- G-Q1: Q1 solo cambia en que la ausencia banco deja de abortar; **cualquier movimiento Q1 fuera de la lista cerrada = STOP con causa byte-probada**.

### S1 — Criterios H0/T0 (read-only, cero código de producción)

- S1.1: matriz headers por (ASE × período) en julio-1, julio-2, agosto-2, octubre-1: headers observados (texto crudo + índice físico como evidencia, no contrato), presentes/ausentes vs `EditablesPorAse[ase].Keys`.
- S1.2: `Total` distinguido de fuente en todas las muestras (exclusión explícita; ninguna fuente real normaliza a `total`).
- S1.3: conceptos 1/2/3/7 + fila `Total` se leen igual (mismas etiquetas/prefijos; fila 7 ausente en Q1 = 0 demostrable intacto).
- S1.4: Σ conceptos == `TotalFuente` por empresa presente en las muestras (±0.5; gate i intacto, sin cambios de código ni de umbral).
- S1.5: veredicto ASE2-octubre1 (observado {ENEL, OCCIDENTE}; `NUEVO ESQUEMA` ausente → omisión + huella) vs ASE4/ASE5 (3 observadas = 3 esperadas) + paridad global esperado↔observado (módulo ausencias legítimas).
- S1.6: inventario de headers no reconocidos en las muestras (esperado: vacío; si aparece alguno → STOP y reporte, nunca clasificación silenciosa).
- S1.7: fixture read-only verde (`ReporteBancoHeadersT0Plan51Tests`, BCL ZIP+XML read-only, cero producción; si requiere celdas numéricas, solo-lectura sin tocar producción — citado) que congela S1.1-S1.6.

### S2 — Criterios H1 (A1 + A2 + verificación A4)

- S2.1: header-scan emite **solo** fuentes reconocidas del vocabulario D-1; la ausente no se exige, no se leen sus conceptos, no se emite su fila.
- S2.2: ausente → omisión + huella (`FuentesAusentesBanco` = `BANCO:{label}` → `ResultadoProcesoPeriodo.FuentesAusentes` como `ASE {n} · BANCO:{label}`) + 1 `Warning` por (ASE, fuente) en formato administrativo (doctrina Plan 26 D-E: sin jerga de código).
- S2.3: repro 2026101-ASE2: el resumen {ENEL, OCCIDENTE, Total} ya **no** lanza el throw posicional; `NUEVO ESQUEMA` queda omitido con huella.
- S2.4: conceptos/fila `Total`/`TotalFuente`/gate (i)/C59/`Quincena` intactos en comportamiento (goldens banco existentes sin deltas fuera de ausencias).
- S2.5: writer verificado sin cambios (test: bloque con ausente escribe presentes + C59; celdas de la ausente intactas en 0) + `ColumnasFuentePorAse` retirada del consumo (borrado u `[Obsolete]` según §7.5, con grep que lo demuestra).
- S2.6: build 0 warnings + greps D-H + G-FORMULA + G-SHA + G-Q1.

### S3 — Criterios H2 (A3 + regresión)

- S3.1: A3-banco implementado (evidencia aditiva del reader + juicio en `ValidadorBasico`): header mutado simulado + ausente en el mismo resumen → fail-fast con formato §2.6; ausencia coherente (subconjunto válido) → verde con huella; desconocida sin ausencia → solo evidencia.
- S3.2: e2e 2026101 + Q1/Q2 de referencia: tabla corrida→nueva justificada por celda; cambios solo trazados a ausencias banco (omisión + huella).
- S3.3: set-diff cumple S0 (altas == lista cerrada citada: solo tests nuevos Plan 51 + bajas citadas si el T0 las exige; hoy **ningún test existente** referencia el throw posicional — grep §0.1: solo producción + `Docs/T0-Plan50…` lo citan) + SHAs intactos + code-review independiente (spec compliance + calidad).
- S3.4: si H0 declaró alguna muestra no absorbible sin ramas → residuo declarado con su T0 propio (patrón Plan 50), nunca parche.

### S4 — Criterios H3 (docs/cierre)

- S4.1: doctrina del plan en `.opencode/project-context.md` (header-scan banco, huella `BANCO:`, A3-banco; solo lo probado).
- S4.2: nota en `plans/09` (T0-0.3 posicional superseded por header-scan de este plan; T0-0.7 de lectura dinámica ratificado) sin reescribir historia.
- S4.3: `Docs/Instructivo-Capa-B.md` si aplica (qué significa `ASE n · BANCO:…` para el administrativo: 0 con huella = normal; etiqueta banco no reconocida = escalar).
- S4.4: diff de docs citado; cero código de producción en H3.

---

## 4. TASKS (fases H0..H3 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero código de producción)

| ID | Tarea | Referencia | Verificación |
|---|---|---|---|
| H0.1 | Matriz de headers por bytes (BCL ZIP+XML read-only, cero producción): por (ASE × período) en julio-1, julio-2, agosto-2, octubre-1 — fila del `Resumen Recaudo Aplicado Por Servicio`, headers observados (texto + índice físico como evidencia) vs `EditablesPorAse[ase].Keys` | `WorkbookLeafCellMapReporteBanco.cs:48-94`; fuentes `Docs/Prueba Octubre-1/Insumos/2-Lime/…`, `Docs/Insumos/` y períodos Q1/Q2 del Ingeniero | Tabla S1.1 publicada |
| H0.2 | Veredicto `Total`-vs-fuente: exclusión explícita por normalizado `total` antes del vocabulario; probar que ninguna fuente real colisiona | `ExcelDataReaderWorkbookLeafInputReader.cs:255-265,290-298` | S1.2 escrito |
| H0.3 | Veredicto conceptos + fila `Total` intactos: mismos prefijos/etiquetas en las 4 muestras; fila 7 ausente Q1 = 0 demostrable | `WorkbookLeafCellMapReporteBanco.cs:35-41`; `ReporteBancoInputs.cs:34-38` | S1.3 escrito |
| H0.4 | Veredicto gate: Σ conceptos == `TotalFuente` por empresa presente en las muestras (±0.5); gate (i) sin cambios de código ni umbral | `WorkbookLeafCoherence.cs:139-153`; `ValidadorBasico.cs:737-767` | S1.4 escrito |
| H0.5 | Veredicto ASE2 (2 fuentes, `NUEVO ESQUEMA` ausente en octubre-1) vs ASE4/ASE5 (3) + paridad global esperado↔observado | `EditablesPorAse` (`:63-94`) | S1.5 escrito |
| H0.6 | Inventario de headers no reconocidos en las muestras (vocabulario D-1); si aparece alguno → STOP y reporte | `EmpresaFacturacion.cs:54-111`; `SinonimosEmpresaR1.cs:38-46` | S1.6 (vacío esperado o STOP) |
| H0.7 | Congelar fixture read-only nuevo `ReporteBancoHeadersT0Plan51Tests` (gemelo de `FuentesRecaudoOpcionalesT0Plan50Tests`) con S1.1-S1.6 | `Remuneracion.IntegrationTests/` | Fixture verde; S1.7 |

- **Puerta H0→H1:** fixture verde + veredictos S1.1-S1.6. Si H0 demuestra forma no absorbible sin ramas → STOP declarado. Sin T0 verde no hay H1.

### H1/F1 — A1 + A2 + verificación A4 (solo lo probado por el T0)

| ID | Tarea | Archivo:línea | Verificación |
|---|---|---|---|
| H1.1 | Header-scan en `LeerReporteBanco`: reemplaza el loop posicional `:255-265` por recorrido de la fila de headers con vocabulario D-1; esperadas-no-observadas → omisión + huella; observadas-no-reconocidas (≠`Total`) → evidencia; cero reconocidas → fail-fast | `ExcelDataReaderWorkbookLeafInputReader.cs:247-265` | S2.1; repro 2026101-ASE2 sin throw posicional (S2.3) |
| H1.2 | Campos aditivos `FuentesAusentesBanco` + `EtiquetasBancoNoReconocidas` (`BANCO:{label}`) en `ReporteBancoAseInputs` + propagación aditiva en `ProcesadorPeriodo` (`:560-569`) a `FuentesAusentes`/`EtiquetasNoReconocidas` como `ASE {n} · BANCO:{label}` + 1 `Warning` por (ASE, fuente) | `ReporteBancoInputs.cs`; `ResultadoProcesoPeriodo.cs:44-58`; `ProcesadorPeriodo.cs:550-570` | S2.2; vacío = comportamiento anterior |
| H1.3 | Verificación A4 con test: bloque con ausente → writer escribe presentes + C59, celdas de la ausente intactas en 0; **cero código** en `OpenXmlPlantillaWriter.cs` salvo T0 en contra | `OpenXmlPlantillaWriter.cs:1329-1359` | S2.5 (primera parte) |
| H1.4 | Retirar el consumo de `ColumnasFuentePorAse` (destino §2.5/§7.5: borrado tras paridad T0 u `[Obsolete]` citado) + grep de cierre | `WorkbookLeafCellMapReporteBanco.cs:48-56` | S2.5 (segunda parte) |
| H1.5 | Tests focalizados (rojo→verde honesto): ausente → omisión + huella; `Total` nunca es fuente/evidencia; desconocida sin ausencia → solo evidencia; cero reconocidas → fail-fast; reconocida-sin-celdas → fail-fast (desync); conceptos/`TotalFuente`/gate intactos | `Remuneracion.IntegrationTests/` (nuevos `ReporteBancoFuentesOpcionalesPlan51Tests.cs`) | S2.1-S2.4; build 0 warnings; greps D-H |
| H1.6 | Regla de reconocida-sin-destino: conservar el fail-fast vigente (`:1344-1348` o gemelo en reader) con formato ASE + `BANCO` + empresa | `OpenXmlPlantillaWriter.cs:1342-1348` | Test que la demuestra; nunca skip silencioso |

- **Puerta H1→H2:** S2.1-S2.6 verdes. Sin esto no se abre H2.

### H2/F2 — A3 + regresión

| ID | Tarea | Archivo:línea | Verificación |
|---|---|---|---|
| H2.1 | Juicio A3-banco en `ValidadorBasico` (método nuevo junto a `ValidarFuentesRecaudoNoReconocidas`, `:891-933`; mensaje §2.6/§3-S0 G-TRAZA); `ClasificarDivergenciasDeclaradas` lo hereda (última puerta) | `ValidadorBasico.cs` | S3.1 (juicio) |
| H2.2 | Test de simulación A3 (sintético donde sea posible + e2e con fuente mutada donde el proyecto lo exige): etiqueta mutada + ausente → fail-fast con Δ/etiquetas exactas; ausencia coherente → verde con huella | `Remuneracion.IntegrationTests/` (nuevo `ControlBancoNoReconocidoPlan51Tests.cs`) | S3.1; 0 `Skip` |
| H2.3 | Regresión e2e 2026101 (+ Q1/Q2 de referencia): tabla corrida→nueva justificada por celda; suite trx + set-diff vs base `p50-h2-suite.trx` (**1141/1074/67/0**) con lista cerrada; SHAs; G-Q1 | `Remuneracion.IntegrationTests/TestResults/p51-h2-suite.trx` + `p51-h2-setdiff.txt` | S3.2-S3.3 |
| H2.4 | Code-review independiente (spec compliance + calidad): trazabilidad cada cambio → A1/A2/A3/A4; ningún test debilitado; CRLF; SHAs | — | Puerta H2→H3: review APROBADO (o APROBADO-CON-OBSERVACIONES cerradas) |

- **Puerta H2→H3:** set-diff S0 + G-Q1 + SHAs + review aprobado.

### H3/F3 — Docs/cierre

| ID | Tarea | Archivo | Verificación |
|---|---|---|---|
| H3.1 | Doctrina en `.opencode/project-context.md` (header-scan banco + huella `BANCO:` + A3-banco; solo lo probado) | `.opencode/project-context.md` | Diff citado; S4.1 |
| H3.2 | Nota en `plans/09` (T0-0.3 posicional superseded por este plan; T0-0.7 ratificado) sin reescribir historia; desviaciones con fecha/motivo/evidencia | `plans/09 - HU-09 Reporte recaudo por banco Fase 2.md` | Diff citado; S4.2 |
| H3.3 | `Docs/Instructivo-Capa-B.md` si aplica (`ASE n · BANCO:…` = normal; etiqueta banco no reconocida = escalar) | `Docs/Instructivo-Capa-B.md` | S4.3 |

- **Puerta H3:** diff de docs citado; cero código de producción en H3.

**Orden sugerido:** H0 → H1 → H2 → H3. Regresión continua (suite + SHAs + G-Q1 en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

---

## 5. Riesgos (tabla consolidada + gates por fase)

Ver §1.4 (R-FALSO-0, R-TOTAL, R-VACIO, R-MAPA-DESYNC, R-FIXTURE, R-RUIDO, R-Q1, R-CANCEL, R-ALCANCE). Gates por fase:

- **Gate H0→H1:** fixture T0 verde + veredictos S1.1-S1.6. Sin T0 verde no hay H1; si H0 declara STOP, no se abre H1.
- **Gate H1→H2:** S2 verde (solo ausencias probadas) + build 0w + greps D-H + G-FORMULA + G-SHA + G-Q1.
- **Gate H2→H3:** set-diff S0 (altas == lista cerrada) + Q1 intacto + SHAs intactos + tabla e2e justificada + review aprobado.
- **Gate H3:** diff de docs citado; cero código de producción en H3.

---

## 6. Architecture Validation Certificate

- **SOLID:** SRP (reader = evidencia + omisión + huella / validador = juicio A3 + gates / mapas = datos sin índices físicos; writer no juzga); OCP (fuente ausente o nueva = cero código; parcial nuevo = fila citada en T0, cero ramas); LSP/DIP (dependencias por interfaces existentes; campos aditivos, contratos intactos); ISP (contratos `IWorkbookLeafInputReader`/`IValidador`/`IPlantillaWriter` intactos).
- **Best practices:** contract-first (H0 congela headers + ausencias + reglas antes del juicio); fail-fast accionable (ASE + `BANCO` + fuente/etiqueta); gates que leen lo escrito + dominio (nunca `<v>` stale); tolerancia ±0.5; values-only; D-H; CRLF; sin commits.
- **Performance:** O(headers + filas ya leídas) + O(ASE × header) en memoria; sin workbooks extra; sin re-lectura por celda; regens a temp fresca.

---

## 7. Clarification Gate (requiere aprobación del Ingeniero; sin esto no se abre H1)

1. **D1-A (vocabulario global + esperado = `EditablesPorAse[ase].Keys`)** frente a D1-B (lista propia del banco) y D1-C (solo catálogo): ¿se ratifica no crear ningún catálogo nuevo?
2. **D2-A (campo propio `FuentesAusentesBanco` → `ResultadoProcesoPeriodo.FuentesAusentes` como `ASE n · BANCO:{label}`)** frente a D2-B (reutilizar conciliación) y D2-C (campo propio en el resultado): ¿se ratifica el canal durable existente?
3. **Omisión (no emitir filas de la ausente) frente a W-A-cero-explícito del Plan 50:** ¿se ratifica la diferencia (§2.5) dado que `EscribirCeldasBanco` itera presentes?
4. **A3-banco por etiquetas sin umbral (§2.3, reglas 1-6: `Total` excluido, vacías ignoradas, cero-reconocidas → fail, desync → fail):** ¿se ratifica frente a un control aritmético?
5. **`ColumnasFuentePorAse`: borrar tras paridad T0** frente a conservar `[Obsolete]`/fallback: ¿se ratifica el borrado?
6. **Puerta:** base `p50-h2-suite.trx` (**1141/1074/67/0**) para el set-diff; rojo fuera de lista = STOP; sin 1–5 no se abre H1; sin H1 verde no hay H2; sin review no hay H3.

Sin estas aprobaciones el plan no avanza de H0 a H1 (1, 2, 4, 5) ni cierra H2 (toda la regresión + §S3) ni H3 (6).

**RATIFICACION 2026-10-22 (Ingeniero) — APROBADO 6/6:** (1) D1-A: vocabulario global + esperado = `EditablesPorAse[ase].Keys`, sin catálogo nuevo; (2) D2-A: huella `BANCO:{label}` en `FuentesAusentes` existente; (3) omisión (la celda queda en el 0 de la plantilla), no cero explícito; (4) A3-banco por etiquetas sin umbral, con `Total` excluido y vacío/cero-reconocidas/desync → fail; (5) borrar `ColumnasFuentePorAse` tras paridad T0; (6) base `p50-h2-suite.trx` (1141/1074/67/0), rojo fuera de lista = STOP. **H0 HABILITADO.**

---

## 8. Estado

**APROBADO 2026-10-22.** Clarification Gate §7 ratificado 6/6 por el Ingeniero. **H0 habilitado:** T0 read-only bloqueante (matriz de headers por ASE × período, `Total`, conceptos, gate Σ y unknowns). Sin H0 verde no se abre H1.

### Estado de implementación (2026-10-22, H0→H3)

| Fase | Estado | Evidencia |
|---|---|---|
| **H0** (T0 read-only) | **VERDE sin STOP** | `Docs/T0-Plan51-Headers-Banco.md` (matriz 20 muestras con fila real de Excel, §1-§6) · fixture `ReporteBancoHeadersT0Plan51Tests` (105 verdes) |
| **H1** A1 (header-scan) | **DONE** | `ExcelDataReaderWorkbookLeafInputReader.LeerReporteBanco` (header-scan + `EsFuenteDelVocabulario`) · `ColumnasFuentePorAse` **borrado** de `WorkbookLeafCellMapReporteBanco` · repro 2026101 ASE2 sin throw posicional |
| **H1** A2 (huella) | **DONE** | `ReporteBancoAseInputs.FuentesAusentesBanco` (`BANCO:{label}`) → `ResultadoProcesoPeriodo.FuentesAusentes` (`ASE n · BANCO:{label}`) + 1 `Warning` por (ASE, fuente) |
| **H1** A4 (writer) | **VERIFICADO sin cambios** | `EscribirCeldasBanco` intacto; test `Writer_BloqueConFuenteAusente_EscribePresentesYC59_DejaLaAusenteEnCero` |
| **H2** A3-banco (juicio) | **DONE** | `ValidadorBasico.ValidarEtiquetasBancoNoReconocidas` (etiquetas, sin umbral) · `ControlBancoNoReconocidoPlan51Tests` (4 verdes) |
| **H2** regresión | **VERDE** | `TestResults/p51-h2-suite.trx` = **1260/1193/67/0** vs base `p50-h2-suite.trx` = 1141/1074/67/0 → **119 altas (solo tests de este plan), 0 bajas, 0 rojos nuevos** (`TestResults/p51-h2-setdiff.txt`) |
| **H3** docs | **DONE** | `.opencode/project-context.md` (doctrina Plan 51) · `plans/09` (nota: T0-0.3 superseded, T0-0.7 ratificado) · `Docs/Instructivo-Capa-B.md` §12 |

**Desviación declarada del T0 (sin STOP):** el §3-S1.5 esperaba que sólo ASE2 llegara con 2 fuentes en octubre-1. La
medición de bytes mostró que **ASE4 de octubre-1 también** trae `{ENEL, OCCIDENTE}` (sin `NUEVO ESQUEMA`). Es la misma
familia de ausencia legítima (label completo, observado ⊆ esperado, cero headers desconocidos) y se absorbe con la regla
ratificada sin ramas; sus celdas `F41:F44` quedan en el 0 de la plantilla y la huella es `ASE 4 · BANCO:NUEVO ESQUEMA`.
Evidencia: `Docs/T0-Plan51-Headers-Banco.md` §5.1 y test `T0P51_Octubre1_Ase4_TambienOmiteNuevoEsquema_VeredictoMedido`.

**Pendiente de la puerta H2→H3:** code review independiente (H2.4) — lo ejecuta el orquestador (el implementer no puede
ser su propio revisor). El resto de gates (build 0/0, set-diff, SHAs, CRLF, G-FORMULA, G-D-H) está verde.

### Cierre (2026-10-22)

- **Code-review independiente:** **APROBADO CON OBSERVACIONES** (0 bloqueantes). Build 0 warnings re-ejecutado; trx re-contado por XML (1260/1193/67/0); 119 altas del plan, 0 bajas, 67 rojos idénticos nombre a nombre; el rojo `GoldenReporteBancoTests.ProcesadorPeriodo_CapaA2_...` es preexistente por `ERR-SIN-SLOT` Q1 (Planes 41/46), no regresión del 51.
- **W-1 corregido:** `ReporteBancoFuentesOpcionalesPlan51Tests.cs` tenía comentarios doble-codificados (`sÃ³lo`, `â†’`, `â”€`); se reescribieron a UTF-8 real (0 restos, CRLF 384/384, build 0 warnings).
- **S-1/S-2 declaradas (no bloquean):** granularidad leaf vs bloque-ASE en `ValidarEtiquetasBancoNoReconocidas` (hoy equivalente, 1 bloque/leaf) y cobertura dual H1/H2 del caso desconocida+ausente.
- **Estado: CERRADO 2026-10-22.** A1+A2+A3+A4 verdes; `ColumnasFuentePorAse` retirado; writer y plantillas intactos. El e2e 2026101 sigue con bloqueos ajenos declarados (`R-PARCIAL-OCT-ASE3-R1-OCC`, `R-SIGMA-Q2`, `ERR-SIN-SLOT` Q1, `R-MAPA-V2-R2-ASE4`, 20 rojos `d55e621`).

