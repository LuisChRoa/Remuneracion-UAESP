---
titulo: Plan 64 - ANT EXT-REV desde RecaudosReversados (escritura en valores por ASE, multi-ASE)
estado: APROBADO 2026-10-30 + ENMIENDA M1–M6 2026-10-30 (autorizada por el Ingeniero tras STOP corto de H0) — Clarification Gate §0.5 ratificado; H0–H4 verdes; H5 EJECUTADO 2026-10-30 (REQ-10 + CA-06 parcial: review independiente PENDIENTE, ver §8-Desviaciones)
fecha: 2026-10-30 (enmienda 2026-10-30)
numero: 64
planes_ocupados: 01..63 ocupados; este es el 64
padre: brief del orquestador 2026-10-30 (funcionalidad ANT EXT-REV) + `Docs/Detalle de plantilla.docx` (sección ANT EXT-REV / ANTICIPOS USUARIOS, rector)
madre_doctrina: .opencode/project-context.md (plantilla autoridad; values-only; guard anti-fórmula; tolerancia ±0,5; prohibido <v> stale como oráculo; 0 warnings; D-H cero quemados; Q1/Q2 por dominio) + plans/59 (geometría derivada, candado anti-quemado, v4 vigente) + plans/57 (anclas derivadas, fin de tabla congelada) + plans/61 (entrega con novedades, columna intermitente por encabezado, doctrina de fixtures) + plans/62 (pegado R10, verificación de integridad read-back) + plans/51 (fuentes opcionales por header + huella + A3 por etiquetas) + plans/26 (preflight, superficie = runtime resuelve) + plans/27 (firma binaria + hoja por nombre) + plans/31-T4 (veredicto R-D-7) + plans/34 (sinonimia NUEVO ESQUEMA ↔ RECIPROCIDAD)
caso_rojo: no hay rojo de corrida — hay AUSENCIA de escritura. La app NO escribe la hoja `ANT EXT-REV` (sus 130 celdas con `<f>` están fuera de destino por el guard; sin `<sheetProtection>` — M3; el comparador la excluye) y NO consume el insumo `RecaudosReversados_*.xlsx` (ni finder ni lectura ni preflight). El manual del administrativo SÍ la trae diligenciada en valores por ASE desde ese insumo.
gate_entrada:
  plantilla_v4: Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx — VIGENTE. Huella semántica (helper canónico `HuellaSemanticaPlantilla`) a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55 (Plan 59 §0.6). Huella distinta ⇒ STOP.
  suite_base: RATIFICADO por enmienda M1–M6 (T0 §7.2–§7.3): base 1799 (1776 verdes + 23 rojos de entrada ajenos a H0, lista cerrada en T0 §7.3) → H0 1817 (+18 verdes, 0 rojos nuevos, 0 desaparecidos). G-SUITE = set-diff exacto contra ese baseline; rojo nuevo fuera de lista = STOP
  sin_commits: NUNCA git commit / git push; NUNCA git checkout / git reset
  excel_cerrado: Excel debe estar cerrado durante las corridas y los T0 por bytes
---

# Plan 64 — ANT EXT-REV desde RecaudosReversados (escritura en valores por ASE)

> SDD completo PROPOSE → DESIGN → SPEC → TASKS → GATES. Fuente única de verdad para la implementación una vez aprobado ESTE documento. APROBADO por el Ingeniero + ENMIENDA M1–M6 2026-10-30 (ver §0.6); H1 habilitado bajo esa enmienda.
> Lectura obligatoria previa: `.opencode/project-context.md` (doctrina madre) + `plans/59` (geometría derivada + candado) + `plans/57` (anclas derivadas) + `plans/61` (Especiales por encabezado + entrega con novedades) + `plans/62` (pegado R10 + integridad read-back) + `plans/51` (fuentes opcionales + huella + A3) + `plans/26` (preflight) + `plans/27` (firma + hoja por nombre) + `plans/31-T4-Veredictos-R2R4-soporte.md` (R-D-7) + `Docs/Detalle de plantilla.docx` (sección ANT EXT-REV / ANTICIPOS USUARIOS, rector).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/WinForms/Cli) + Excel; persistencia por archivos Excel, NO base de datos. Español técnico; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan; WinForms/CLI solo heredan estado y código de salida vigentes).

---

## §0. Estado, rector, alcance, decisiones y riesgos

### §0.1 Estado y rector

- **Estado:** APROBADO + ENMIENDA M1–M6 (2026-10-30). H0 cerrado en verde (18/18 + evidencia `Docs/T0-Plan64-Ant-Ext-Rev.md`); puerta H0→H1 ABIERTA condicionada a esta enmienda (ver §0.6). Sin candados ni review no hay cierre.
- **Rector (IRRENUNCIABLE):** `Docs/Detalle de plantilla.docx`, sección ANT EXT-REV / ANTICIPOS USUARIOS. Ante conflicto entre este plan y el docx, **manda el docx** y el plan se enmienda (el H0 lo detecta por bytes; divergencia docx-vs-encargo = hallazgo, ver D-10).
- **Alcance de flujo:** **multi-ASE (`ProcesadorPeriodo.Ejecutar`) en AMBAS quincenas.** Single-ASE (`ProcesadorRemuneracion`) fuera de alcance por decisión D-01 (el preflight single-ASE solo exige R1/R2/R4 — Plan 26 §V7; extenderlo es plan hijo con su T0).

### §0.2 Alcance (in / out)

**Entra:**

| ID | Entregable | Criterio de salida |
|---|---|---|
| H0/T0 | Inventario read-only por bytes: geometría ANT EXT-REV v4, fórmulas del bloque superior y controles, celdas de valor, 100 %-fórmulas de ANTICIPOS USUARIOS, referencias L125:L129 + L130 vacía + L131 consumidora (M6), estructura Sheet1 de `RecaudosReversados_*`, universo de fuentes en 30 archivos, paridad numérica citada | Fixture verde + `Docs/T0-Plan64-Ant-Ext-Rev.md`; puerta H0→H1 |
| H1 | Lector `LeerAntExtRev` + modelos Core + catálogo de fuentes + geometría derivada por etiquetas + validación estructural fail-fast | CA de §3 en lectura pura (sin escribir) |
| H2 | Escritor values-only (100 celdas) en la pasada atómica + read-back ±0,5 + fórmulas protegidas intactas | CA de §3 en corrida 2026091 y 2026101 |
| H3 | Preflight (`BuscarRecaudosReversados` + `ReporteInsumoAse.RecaudosReversados` + firma + `FormateadorInsumosFaltantes` + `ValidadorInsumosPeriodo` en ambas quincenas) + actualización de doctrina §V9 y comentarios mentirosos | CA de §3 en período incompleto (fail-fast agregado) |
| H4 | Comparador: retirar exclusión ANT EXT-REV, actualizar BrechaDeclarada ítem 8, goldens por bloque/ASE contra los 4 manuales | CA de §3 en comparador |
| H5 | Prevención sistémica (candado anti-quemado + sonda de etiquetas + doctrina de fixtures) + cierre | Review APROBADO + suite set-diff exacto |

**No entra (vinculante):**

1. Single-ASE (`ProcesadorRemuneracion`): no lee ni escribe ANT EXT-REV; su guardrail sigue siendo R1/R2/R4.
2. `ANTICIPOS USUARIOS` y `CONSOLIDADO_TOTAL RECAUDO!L125:L131` (M6: L125:L129 referencian, L130 vacía, L131 consumidora): 100 % fórmulas que consumen lo escrito; **no se escriben** (recalculan solas; `fullCalcOnLoad` vigente, Plan 28).
3. Nuevos mapeos fuente→bloque fuera del catálogo T0 (EAAB CIUDAD LIMPIA y ENERBIT quedan en 0 por ausencia; ver D-10 y hueco HUE-1).
4. Cambios a la plantilla (huella v4 intacta; cero `<f>` nuevo, cero `InsertRow`/`DeleteRow`).
5. Cambios de negocio en fórmulas fuera de la resolución de geometría por etiquetas.
6. Commits/push/checkout/reset. NUNCA `git commit` / `git push`. NUNCA `git checkout` / `git reset`.

### §0.3 Decisiones ratificables D-01..D-10 (el Ingeniero las ratifica o mejora en §0.5)

| ID | Decisión propuesta | Fundamento | Alternativa descartada |
|---|---|---|---|
| D-01 | Alcance multi-ASE (`ProcesadorPeriodo`), ambas quincenas; single-ASE fuera | El preflight single-ASE es guardrail mínimo R1/R2/R4 (Plan 26 §V7); la hoja es agregación por ASE del período y su insumo es por ASE. Q1/Q2 por dominio (G2-D2) | Meterlo en single-ASE (rompería la doctrina del guardrail mínimo sin T0 propio) |
| D-02 | Extender `IWorkbookLeafInputReader` con `LeerAntExtRev` implementado en `partial` nuevo del reader | **Corrección al encargo:** no son "25 composiciones" — el grep mide **100+ sitios** `new ExcelDataReaderWorkbookLeafInputReader()` (truncado en 100; ver §1.6). Una interfaz dedicada obligaría a retocar todos los implementadores/mocks/composiciones. `partial` nuevo = cero toque a los 100+ sitios; el contrato crece aditivamente | Interfaz dedicada `IAntExtRevReader` (costo de composición y doble localización del mismo insumo; se reconsidera solo si H1 demuestra acoplamiento real) |
| D-03 | Modelos Core `AntExtRevAseInputs` + catálogo de fuentes con etiqueta de bloque y citas + `WorkbookLeafInputs.AntExtRev` nullable | Patrón Planes 34/42/51: tabla citada + matcher puro; `null` = comportamiento anterior intacto (compat por construcción). Sin cita no hay fila (R-SINONIMO-ABUSO) | Campos sueltos por fuente en `WorkbookLeafInputs` (quema la forma; la próxima fuente exige código nuevo) |
| D-04 | Escribir SOLO fuentes presentes (ausente = cero del template + traza tipo Plan 51); los 4 conceptos de una fuente presente con 0 explícito si falta el concepto | Doctrina Plan 50/51: la fuente ausente es caso normal (0 + huella); el concepto ausente dentro de fuente presente es 0 explícito porque el writer itera presentes y la celda existe. Prohibido inventar filas. **Adicional H1-manual (T0 §5.3/§8-H1): `2026092!J20` VACÍA en el manual donde esta regla manda 0 ⇒ el golden H2/H4 trata vacío ≡ 0** | Cero explícito también para fuentes ausentes (inventaría `TotalFuente=0` indistinguible de dato real; el Plan 51 §2.5 ya resolvió omisión-vs-cero a favor de omisión) |
| D-05 | Geometría derivada de la plantilla por etiquetas con validación estructural y fail-fast `ERR-PLANTILLA`; **prohibido hardcodear C11..L12**; tabla T0 congelada como evidencia de test | Doctrina Plan 59 §0.2: cero literales de fila/celda/ventana en producción. La T0 es EVIDENCIA del fixture, nunca contrato de producción | Re-congelar C11..L12 (M-B del Plan 59: la próxima plantilla rompe igual) |
| D-06 | **(ENMENDADO M4 — BLOQUEANTE.)** Fail-fast `ERR-FORMATO-FUENTE` si: faltan/invierten marcadores, no hay fuentes, etiqueta desconocida o duplicada, **D≠E≠F > 0,5 (se exige igualdad D=E=F; se lee F como valor del concepto)**, Σ grupo ≠ total fuente > 0,5. **"Por encabezado" SOLO rige para C (concepto) y etiquetas de canal/fuente: NO hay encabezados semánticos para D/E/F** (T0 §4.1: D='Cantidad' pertenece a la tabla de pendientes; cabecera real E='ASEO', F='Total', D='66'). **PROHIBIDO implementar D/E/F "por índice" como mapa quemado — la igualdad es la invariante, no el índice** | La hoja es dinero por fuente; una estructura rota o una aritmética que no cierra es defecto interno (nunca declarable, Plan 42). Mensaje con ASE + reporte + marcador/etiqueta + esperado vs obtenido (en D/E/F: ASE + fuente + concepto + D/E/F observados) | Warning-y-continuar (viola atomicidad fail-fast; entregaría dinero parcial como si fuera total) |
| D-07 | Preflight: `BuscarRecaudosReversados` + prefijo normalizado `recaudosreversados` + `ReporteInsumoAse.RecaudosReversados` + `FormateadorInsumosFaltantes` + `ValidadorInsumosPeriodo` en AMBAS quincenas; actualizar doctrina §V9 y comentarios mentirosos | Superficie = runtime resuelve (Plan 26 D-D): al consumir el insumo, el preflight DEBE exigirlo. Q1/Q2 por dominio (el insumo aplica en ambas). Los comentarios "§V9 / fuera" quedan falsos y son parte del defecto (causa tipo Plan 61 §0.2-c) | Exigirlo solo en Q2 (inventa una regla de negocio sin cita; el encargo ordena ambas quincenas) |
| D-08 | Gates: read-back ±0,5 + fórmulas protegidas intactas + coherencia internos | Lección Plan 31: el gate lee el WORKBOOK generado (texto-`<f>` + literales), no solo el dominio. Prohibido `<v>` stale (Planes 28/29/40) | Gate dominio-a-dominio (cancelaría una pérdida común, Plan 49) |
| D-09 | Comparador: retirar exclusión ANT EXT-REV, actualizar BrechaDeclarada ítem 8, goldens por bloque/ASE contra los 4 manuales | Al escribir la hoja, la exclusión pierde su causa (non-goal del Plan 29). Retiro versionado con cita T0, no blanket (Plan 29 D-E) | Conservar la exclusión (falso verde permanente sobre dinero escrito) |
| D-10 | Divergencia declarada docx-vs-encargo: el docx rector dice usar el "Resumen" y la explicación administrativa exige el bloque 2 para discriminar fuentes → **bloque 2 = fuente, bloque 3 = control**. HUECO NO PROBADO HUE-1: mapeo de EAAB CIUDAD LIMPIA y ENERBIT (0 por ausencia; fail-fast ante etiqueta no tabulada) | Lo no probado se declara hueco (regla del encargo). El bloque 3 no discrimina por fuente (columnas variables; ASE2 Sept-1 no trae OCCIDENTE) ⇒ no puede ser fuente sin inventar reparto | Usar el Resumen como fuente (no discrimina OCCIDENTE en ASE2 Sept-1; repartir sería inventar dinero) o inventar mapeo para CIUDAD LIMPIA/ENERBIT (sin evidencia en 30 archivos ni en 4 manuales) |

### §0.4 Riesgos R-01..R-10 (detalle y gates en §7)

| ID | Riesgo | Mitigación + gate que lo detecta |
|---|---|---|
| R-01 | Re-congelar C11..L12 / B8..B40 (M-B del Plan 59) | D-05 + candado H5 (estático + injerto en rojo) + grep G-QUEMADO |
| R-02 | Letra/columna física del insumo (D/E/F) fijada por índice en vez de por invariante | D-06 enmendado (M4): igualdad D=E=F + lectura de F; test "columna movida" (H1.5) |
| R-03 | Falso verde por `<v>` stale o por gate dominio-a-dominio | D-08 + G-WORKBOOK (texto-`<f>` + literales) |
| R-04 | Fuente ausente tratada como error (o al revés: etiqueta desconocida tratada como ausencia) | D-04 + D-06 + A3 gemelo del Plan 51 (desconocida + ausente ⇒ fail-fast) |
| R-05 | Σ-grupo ≠ total-fuente tolerado en silencio | D-06 (fail-fast > 0,5 con ASE + fuente + esperado vs obtenido) |
| R-06 | Preflight y runtime divergen (nuevo prefijo que el runtime resuelve y el preflight no exige) | D-07 + checklist de paridad preflight↔runtime (Plan 26 R-F-6) |
| R-07 | Comparador sigue excluyendo la hoja (falso verde sobre dinero) | D-09 + test negativo (injerto: divergencia ANT EXT-REV sin exclusión ⇒ rojo) |
| R-08 | Mapeo inventado para CIUDAD LIMPIA / ENERBIT | D-10 + HUE-1 (fail-fast ante etiqueta no tabulada, nunca 0 silencioso si VIENE) |
| R-09 | Doble escritura / orden que deja estados intermedios | §2.8 (orden en la pasada atómica) + grep de cierre |
| R-10 | Comentarios §V9 mentirosos sobreviven al cambio | D-07 (actualización en el mismo cambio; grep de cierre los exige) |

### §0.5 Clarification Gate (requiere aprobación del Ingeniero; sin esto no se abre H0)

1. **D-01 (alcance):** ¿multi-ASE ambas quincenas; single-ASE fuera?
2. **D-02 (contrato):** ¿extender `IWorkbookLeafInputReader` con `LeerAntExtRev` en `partial` nuevo (corrección: no 25 sino 100+ composiciones) frente a interfaz dedicada?
3. **D-03/D-04 (modelos y opcionalidad):** ¿catálogo citado + `AntExtRev` nullable; ausente = cero del template + traza; concepto ausente en fuente presente = 0 explícito?
4. **D-05 (geometría):** ¿derivada por etiquetas con `ERR-PLANTILLA`; prohibido hardcodear C11..L12; T0 como evidencia?
5. **D-06 (fail-fast):** ¿las 6 condiciones `ERR-FORMATO-FUENTE` con el mensaje propuesto?
6. **D-07 (preflight):** ¿`BuscarRecaudosReversados` + `recaudosreversados` + enum + formateador + validador en ambas quincenas + reescritura de §V9?
7. **D-08/D-09 (gates y comparador):** ¿read-back ±0,5 + protegidas intactas + coherencia; retirar exclusión y actualizar ítem 8 con goldens contra los 4 manuales?
8. **D-10 (divergencia y hueco):** ¿bloque 2 = fuente + bloque 3 = control; HUE-1 declarado (CIUDAD LIMPIA/ENERBIT en 0 por ausencia, fail-fast si vienen)?

Sin 1–8 no se abre H0; sin H0 verde no hay H1; sin candados ni review no hay cierre. (Ratificado 2026-10-30; H0 cerrado 18/18.)

### §0.6 ENMIENDA M1–M6 (2026-10-30, autorizada por el Ingeniero tras el STOP corto de H0)

Fuente: `Remuneracion.IntegrationTests/AntExtRevT0Plan64Tests.cs` (18/18 verde) + `Docs/T0-Plan64-Ant-Ext-Rev.md` (§8 trae M1–M6 y baseline; §7.2–§7.3 el set-diff). La enmienda NO cambia alcance ni decisiones D-01..D-10: corrige texto del §1/§2 contra bytes medidos. M4 es BLOQUEANTE (reescribe regla de lectura).

| ID | §1 decía | Byte medido (T0) | Enmienda aplicada en este plan |
|---|---|---|---|
| M1 | §1.1: celdas de valor "texto '0' o vacías" | 100/100 numéricas `<v>0</v>`, sin `t`, sin `<f>` (T0.1d) | §1.1/§2.5/§2.6: "numéricas en 0"; H2 escribe números (no strings), read-back decimal |
| M2 | §1.1: fila+2 REPORTE 3/4 "C..L" | Alterna en **C..N** (12 cols); **C..L** es la rejilla de valores (T0.1a) | §1.1/§2.5: rejilla C..L; encabezados C..N |
| M3 | §1.1: hoja "100 % protegida" | Sin `<sheetProtection>`; "protegida" = 130 celdas con `<f>` fuera de destino (T0 §2.5) | §1.1/§2.5: sin protección de hoja; guard = celdas con `<f>` no destino |
| M4 | §2.4 paso 3 + REQ-04 + D-06: "D/E/F por ENCABEZADO" | Sin encabezados semánticos para D/E/F (D='Cantidad' es de la tabla de pendientes; cabecera real E='ASEO', F='Total', D='66'). Única propiedad estable: **D=E=F ±0,5 en 30/30** (T0 §4.1) | REGLA REESCRITA (ver D-06 enmendado + §2.4 paso 3): exigir igualdad D=E=F con fail-fast si >0,5; leer **F como valor del concepto**; "por encabezado" solo vale para C (concepto) y etiquetas de canal/fuente. PROHIBIDO implementar D/E/F "por índice" como mapa quemado — la igualdad es la invariante, no el índice. HUE-3 CERRADO: ninguna columna "manda" |
| M5 | §1.3: "ASE2 Sept-1 no trae OCCIDENTE" (caso único) | **8 archivos** sin columna `RECAUDO OCCIDENTE` en bloque 3, lista congelada T0 §4.3 | §1.3: lista de 8 (2026071 ASE2, 2026082 ASE1, 2026082 ASE3, 2026091 ASE2, 2026092 ASE3, 2026092 ASE5, 2026101 ASE3, 2026101 ASE5) |
| M6 | §1.1: "`CONSOLIDADO!L125:L130` referencia ANTICIPOS USUARIOS" | `L125:L129 = -'ANTICIPOS USUARIOS'!C21..C25`; **`L130` VACÍA**; `L131=SUM(L125:L130)` consumidora (T0.2c) | §1.1/§2.6/CA-03: rango que referencia L125:L129; L131 consumidora |

Adicional H1-manual (T0 §5.3 + §8-H1): `2026092!J20` VACÍA en el manual donde D-04 manda 0 explícito ⇒ el golden H2/H4 trata **vacío ≡ 0** (no inventar valor). Nota de implementación T0 §3.1/M7: `O26` es seguidora de compartida `si=3` — expandir shared al barrer `<f>`.
Baseline ratificado (§5/§6): base 1799 (1776 + 23 rojos de entrada, T0 §7.3) → H0 1817 (+18 verdes, 0 rojos nuevos, 0 desaparecidos). Puerta H0→H1: ABIERTA.

---

## §1. Evidencia T0 (lo citado por el encargo; el H0 lo re-mide — nada de aquí es oráculo)

> Método H0: BCL puro (`ZipArchive` + `X.Linq`) para la plantilla v4 y los 4 manuales; ExcelDataReader solo-lectura sobre COPIAS para los 30 `RecaudosReversados_*`. Rutas explícitas (prohibido `Insumos.Raiz()` en lo nuevo). Cero código de producción en H0: solo fixture `*T0Plan64*` + `Docs/T0-Plan64-Ant-Ext-Rev.md`. Si un número/fórmula/etiqueta de este §1 difiere en disco ⇒ STOP (se re-mide, no se "ajusta" el plan en silencio).

### §1.1 Geometría congelada T0 (v4 y plantilla raíz idénticas — a re-medir)

| Elemento | Estado citado (H0 lo congela) | Verificación en código existente |
|---|---|---|
| Hoja `ANT EXT-REV` | Existe en la plantilla; sin escritura; SIN `<sheetProtection>` (M3: "protegida" = las 130 celdas con `<f>` fuera de destino, guard anti-fórmula) | `WorkbookLeafCellMapQ2.HojaAntExtRev = "ANT EXT-REV"` (`Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapQ2.cs:54`); protegidas `C3/M3/N3/C5` en `:402-405`; gemelo en `WorkbookLeafCellMapAjustesSfT.cs:252-255` |
| Bloque superior filas 3–6 | Fórmulas: C3..N3 = `C11+C19+C27+C35+C43` y variantes; C4..N4 = `C12+C20+C28+C36+C44`; fila 5 SUM; fila 6 auto-chequeo. Controles P3 (`N3=(-VALIDACION_TOTAL!M9)`), P4 (`N4=(-'CONSOLIDADO_TOTAL RECAUDO'!L131)`) | **CITADO por el encargo, NO verificado en código en esta planificación** (los mapas solo protegen C3/M3/N3/C5; el texto-`<f>` completo lo congela el H0 por bytes). Si el texto real difiere ⇒ STOP + enmienda |
| 5 bloques, etiqueta en columna B | B8 BANCO DE OCCIDENTE, B16 EAAB RECIPROCIDAD, B24 EAAB CIUDAD LIMPIA, B32 ENEL, B40 ENERBIT | **CITADO, a congelar en H0** (búsqueda por `B=<etiqueta>`; la etiqueta es DATO leído, nunca literal en producción) |
| Por bloque | fila+1 = ASE1..ASE5 + TOTAL(M); fila+2 = REPORTE 3 / REPORTE 4 alternados **C..N** (M2; **C..L** es la rejilla de valores); fila+3 = APLICADO; fila+4 = ANTICIPOS; fila+5 = TOTAL (SUM por columna) | **CONGELADO por T0.1a** (ver §0.6-M2) |
| Celdas de valor | 10 columnas × 2 filas por bloque = **100 en total**. Naturaleza ENMENDADA M1: **numéricas `<v>0</v>`, sin `t`, sin `<f>`** (100/100, T0.1d); el manual las trae numéricas | **CONGELADO por T0.1d** (naturaleza `<v>`/texto por celda; el writer solo escribe valores numéricos, nunca `<f>`) |
| Protección actual | Solo C3/M3/N3/C5 en ambos mapas | ✅ Verificado: `WorkbookLeafCellMapQ2.cs:402-405` + `WorkbookLeafCellMapAjustesSfT.cs:248-255` |
| Comparador | Excluye la hoja; `BrechaDeclarada` ítem 8 | ✅ Verificado: `Remuneracion.IntegrationTests/ComparadorSalidaVsManualTests.cs:158-162` (`hoja is "ANT EXT-REV" or "ANTICIPOS USUARIOS"` → `"hoja no escrita por la app (non-goal del Plan 29)"`); R-D-7 en `plans/31-T4-Veredictos-R2R4-soporte.md:44` + `.opencode/project-context.md:271` |
| `ANTICIPOS USUARIOS` | 100 % fórmulas, referencia ANT EXT-REV (D4/M4/D20/M20…); `CONSOLIDADO_TOTAL RECAUDO!L125:L129 = -'ANTICIPOS USUARIOS'!C21..C25` (M6); **`L130` VACÍA**; `L131=SUM(L125:L130)` es la consumidora; no se escriben | **CONGELADO por T0.2a–d** (barrido `<f>` por celda + referencias; diff `<f>` antes/después = ∅ es gate). Nota M7: `O26` es seguidora de compartida `si=3` — expandir shared al barrer `<f>` |

### §1.2 Tabla etiqueta→bloque con citas (catálogo propuesto; H0 la congela)

| Etiqueta de bloque (plantilla, col B) | Fuente del insumo (bloque 2, col A) | Cita de negocio | Destino |
|---|---|---|---|
| BANCO DE OCCIDENTE | OCCIDENTE | Identidad por rótulo normalizado (mayúsculas, sin diacríticos — misma regla que `SinonimosEmpresaR1.Normalizar`) | Bloque B8 |
| EAAB RECIPROCIDAD | NUEVO ESQUEMA | `SinonimosEmpresaR1.Clases[0]` = {RECIPROCIDAD, NUEVO ESQUEMA} (`Remuneracion.Core/Models/SinonimosEmpresaR1.cs:38-46`, cita: `Detalle de plantilla` + `Proceso de Recaudo` + oráculo R1!F199) | Bloque B16 |
| ENEL | ENEL | Identidad (catálogo abierto: lo no tabulado solo se iguala a sí mismo, `SinonimosEmpresaR1.cs:51-76`) | Bloque B32 |
| EAAB CIUDAD LIMPIA | — (ausente en 30 archivos; 0 en 4 manuales) | HUE-1: sin evidencia ⇒ sin fila de mapeo (R-SINONIMO-ABUSO: sin cita no hay fila) | Bloque B24 en 0 |
| ENERBIT | — (ausente en 30 archivos; 0 en 4 manuales) | HUE-1: idem | Bloque B40 en 0 |

Concepto→fila/reporte (citado; H0 lo congela con bytes):

| Concepto fuente (col C, normalizado) | Fila destino | Reporte destino |
|---|---|---|
| 1-APLICADOS A FACTURACION | APLICADO | REPORTE 3 |
| 2-SALDOS A FAVOR GENERADOS | ANTICIPOS | REPORTE 3 |
| 9-REVERSION PAGOS A COMPONENTES | APLICADO | REPORTE 4 |
| 9-REVERSION SALDOS A FAVOR | ANTICIPOS | REPORTE 4 |
| ausente en fuente presente | 0 explícito | el de su fila |

Signo: tal cual (los R4 ya son negativos en origen; la plantilla no re-signa).

### §1.3 Tabla de 30 insumos (universo citado; H0 la cierra con conteo)

Universo citado por el encargo (6 períodos × 5 ASE = 30 archivos `RecaudosReversados_*.xlsx`, hoja única `Sheet1`):

- Estructura: marcador `Pagos Pendientes por Aplicar` (fila variable) → **bloque 2** → marcador `Resumen Recaudo Aplicado Por Servicio` → **bloque 3**.
- Bloque 2: filas de concepto con C ∈ {1-APLICADOS A FACTURACION, 2-SALDOS A FAVOR GENERADOS, 9-REVERSION PAGOS A COMPONENTES, 9-REVERSION SALDOS A FAVOR}, valores **D=E=F (M4: tres copias del mismo monto; fail-fast si divergen > 0,5; se lee F)**; filas de total de canal (B=canal, C=Total); fila de total de FUENTE (A=fuente, B=Total) que CIERRA el grupo (etiqueta al final).
- Bloque 3 = RECAUDO OCCIDENTE / RECAUDO OTROS / Total, columnas variables por archivo (**8 archivos sin OCCIDENTE**, lista congelada M5 abajo): **solo control, nunca fuente**.
- Universo de fuentes en 30 archivos: **SOLO ENEL, NUEVO ESQUEMA, OCCIDENTE**. EAAB CIUDAD LIMPIA y ENERBIT no aparecen nunca y están en 0 en los 4 manuales (Julio-2, Septiembre-1, Septiembre-2, Octubre-1).
- Bloque 3 sin OCCIDENTE (M5, lista congelada T0 §4.3 — NO es caso único): `2026071 ASE2`, `2026082 ASE1`, `2026082 ASE3`, `2026091 ASE2`, `2026092 ASE3`, `2026092 ASE5`, `2026101 ASE3`, `2026101 ASE5`.
- H0 congela: conteo cerrado 30/30, marcadores por archivo (fila exacta como evidencia, no como contrato), encabezados D/E/F por archivo, totales de canal y de fuente con su cierre aritmético.

### §1.4 Tabla de paridad por período/ASE/fuente/concepto (citada; H0 la re-deriva)

Verificada numéricamente por el encargo (el H0 la re-deriva `fuente × etiqueta`; mismatch ⇒ STOP):

| Período | ASE | Fuente | Conceptos (APLICADO/R3, ANTICIPOS/R3, APLICADO/R4, ANTICIPOS/R4) | Destino citado |
|---|---|---|---|---|
| Sept-1 | ASE1 | OCCIDENTE (bloque 2) | 42.680 / 40.630 / -83.310 / 0 | → C11, C12, D11, D12 |
| Sept-1 | ASE1 | ENEL R4 | −3.461.197,29 (APLICADO/R4) | → D35 |
| Sept-1 | ASE4 | OCCIDENTE | 239.834 / 1.438.701 / −1.742.535 / 0 | → I11..J12 (fila APLICADO/ANTICIPOS × R3/R4 del bloque BANCO DE OCCIDENTE, col ASE4) |
| Sept-1 | ASE4 | NUEVO ESQUEMA | 208.970 / 53.030 / −262.000 / 0 | → I19..J20 |
| Sept-1 | ASE4 | ENEL | 8.186.910 / 9.321.147 / −20.245.524,78 / −459.462,22 | → I35..J36 |
| Julio-2 | ASE4 | (las 3 fuentes) | coherente con su bloque (cifras en el T0) | → bloque ASE4 por fuente |
| Octubre-1 | ASE4 | (las 3 fuentes) | coherente con su bloque | → bloque ASE4 por fuente |

Invariante citado: **Σ 4 conceptos == total de fuente** (por fuente presente; tolerancia ±0,5).

### §1.5 Verificaciones propias de esta planificación (corrigen el encargo donde difiere)

| # | Encargo decía | Código real (archivo:línea) | Veredicto |
|---|---|---|---|
| V-01 | Comparador excluye la hoja (~línea 159) y BrechaDeclarada ítem 8 | `ComparadorSalidaVsManualTests.cs:158-162` (ítem 8, texto `"hoja no escrita por la app (non-goal del Plan 29)"`); R-D-7 en `plans/31-T4-Veredictos-R2R4-soporte.md:44` y `project-context.md:271` | ✅ Confirmado (línea exacta 159, no aproximada) |
| V-02 | Protección actual: solo C3/M3/N3/C5 en ambos mapas | `WorkbookLeafCellMapQ2.cs:402-405` + `WorkbookLeafCellMapAjustesSfT.cs:252-255` | ✅ Confirmado |
| V-03 | Extender `IWorkbookLeafInputReader` "evita tocar 25 composiciones" | El grep de `new ExcelDataReaderWorkbookLeafInputReader()` devuelve **100+ sitios** (truncado en 100; producción: `ComposicionServicios.cs:42`, `EjecutorCli.cs:183,297`, `Herramientas/VerificadorRecaudo/Program.cs:60` + ~96 en tests) | ⚠️ **Corregido:** no son 25; el argumento a favor del `partial` es AÚN más fuerte (100+ sitios intactos) |
| V-04 | `IWorkbookLeafInputReader` existe con LeerR1/R2/R4 + detalle + banco + balance + saldos + espejo | `Remuneracion.Core/Interfaces/IWorkbookLeafInputReader.cs:8-162` (sin `LeerAntExtRev`; firma `LeerLeafInputs(ase, periodo, rutaR1, rutaR2, rutaR4)` en `:19`) | ✅ Confirmado; `LeerAntExtRev` NO existe (es lo que H1 agrega) |
| V-05 | Sinonimia NUEVO ESQUEMA ↔ EAAB Reciprocidad citada en `SinonimosEmpresaR1` | `SinonimosEmpresaR1.cs:38-46` (clase {RECIPROCIDAD, NUEVO ESQUEMA} con triple cita) + catálogo abierto `:51-76` | ✅ Confirmado |
| V-06 | Preflight: doctrina §V9 + comentarios mentirosos; `ReporteInsumoAse` sin el miembro; sin finder | `ValidadorInsumosPeriodo.cs:16` (`RecaudosReversados queda FUERA`); `ReporteInsumoAse.cs:4-7` (idem); `ILocalizadorArchivosAse.cs:8-82` (sin `BuscarRecaudosReversados`); preflight en `ProcesadorPeriodo.cs:140-155`; escritura en `:402`; gates en `:442+` | ✅ Confirmado; D-07 los actualiza en el mismo cambio |
| V-07 | Geometría fina (C3..N3, P3/P4, B8..B40, 100 celdas texto "0") y paridades numéricas del §1.4 | **CERRADO por T0 18/18 (enmienda M1–M6, §0.6):** 100 celdas numéricas (M1), reportes C..N (M2), sin `<sheetProtection>` (M3), D=E=F sin encabezados semánticos (M4), 8 archivos sin OCCIDENTE (M5), L130 vacía (M6); paridad §1.4 exacta + ampliada a 20 archivos | ⚠️ **Enmendado, ya no pendiente:** el plan NO convierte la T0 en contrato de producción (D-05) |

### §1.6 Huecos (lo no probado se declara)

- **HUE-1 (mapeo EAAB CIUDAD LIMPIA / ENERBIT):** 0 apariciones en 30 archivos + 0 en 4 manuales ⇒ 0 por ausencia con traza. Si VIENEN (etiqueta no tabulada) ⇒ fail-fast `ERR-FORMATO-FUENTE` (nunca 0 silencioso, nunca mapeo inventado). Dueño futuro: T0 de negocio con bytes.
- **HUE-2 (texto-`<f>` exacto del bloque superior y de ANTICIPOS USUARIOS/L125:L130):** citado, a congelar en H0. Si el texto real difiere ⇒ STOP.
- **HUE-3 (columnas D/E/F del bloque 2) — CERRADO por M4 (T0 §4.1 + T0.3/T0.3b):** no hay encabezados semánticos para D/E/F; la regla es **igualdad D=E=F con fail-fast si divergen > 0,5 (D-06) y lectura de F**; ninguna columna "manda". "Por encabezado" solo rige para C y etiquetas de canal/fuente.
- **HUE-4 (Q1 con forma incompatible):** si una fuente Q1 trae forma no cubierta (precedente: secciones Q1 del Plan 41 E1-4) ⇒ `ERR-SIN-SLOT`/`ERR-FORMATO-FUENTE` honesto, nunca fallback posicional.

---

## §2. Diseño

### §2.1 D-H0. Método T0 por bytes (read-only bloqueante, cero código de producción)

Forma normativa: BCL puro (`ZipArchive` + `X.Linq`) para v4 + 4 manuales; ExcelDataReader solo-lectura sobre COPIAS para los 30 insumos. Rutas explícitas a los 6 períodos y a los 4 manuales (Julio-2, Septiembre-1, Septiembre-2, Octubre-1). Prohibido `Insumos.Raiz()` en lo nuevo. Cero código de producción en H0: solo fixture `AntExtRevT0Plan64Tests` + `Docs/T0-Plan64-Ant-Ext-Rev.md`.

Seis frentes: (a) geometría v4 (bloque superior `<f>`, controles P3/P4, 5 etiquetas B, rejilla C..L × filas de valor, naturaleza de las 100 celdas); (b) 100 %-fórmulas de ANTICIPOS USUARIOS + referencias L125:L130; (c) estructura Sheet1 por archivo (marcadores, bloque 2 con encabezados D/E/F, totales de canal/fuente, bloque 3 como control); (d) universo de fuentes 30/30 + paridad numérica §1.4 re-derivada; (e) invariantes Σ-grupo == total-fuente ±0,5; (f) lista cerrada de Δ de suite + huella v4 antes/después.

**Puerta H0→H1:** seis frentes cerrados + fixture verde + APROBADO del Ingeniero. Si algún frente solo cierra con literal/rama/mapeo inventado ⇒ STOP + escalado.

| Candidata | Veredicto | Motivo |
|---|---|---|
| **M-A: T0 de seis frentes + extensión aditiva del contrato + geometría derivada + preflight + gates workbook + comparador versionado** | **Elegida** | Única que escribe dinero real sin quemar geometría ni inventar mapeos; precedente T0-Plan50/51/59/61 |
| M-B: hardcodear C11..L12 / B8..B40 y leer D/E/F por índice | Descartada | M-B del Plan 59: re-fotografía la v4; la próxima plantilla rompe igual o en silencio |
| M-C: interfaz dedicada `IAntExtRevReader` con composición propia | Descartada como principal | Duplica localización del mismo insumo y obliga a retocar 100+ composiciones; solo se reconsidera si H1 demuestra acoplamiento real |
| M-D: usar el bloque 3 (Resumen) como fuente | Descartada | No discrimina por fuente (ASE2 Sept-1 sin OCCIDENTE); repartir sería inventar dinero (D-10) |

### §2.2 Modelos Core (decisión D-03)

```csharp
// Propuesta (nombres exactos a fijar en implementación contra la convención del directorio):
public sealed class AntExtRevAseInputs
{
    public Ase Ase { get; set; } = new();
    // Por fuente PRESENTE: 4 conceptos en el orden (AplicadoR3, AnticiposR3, AplicadoR4, AnticiposR4).
    public IReadOnlyList<AntExtRevFuenteInputs> Fuentes { get; set; } = [];
    public IReadOnlyList<string> FuentesAusentes { get; set; } = [];      // traza tipo Plan 51 ("ANTEXTREV:{fuente}")
    public IReadOnlyList<string> EtiquetasNoReconocidas { get; set; } = []; // evidencia A3 ("ANTEXTREV:{etiqueta}")
}
public sealed record AntExtRevFuenteInputs(
    string FuenteEtiqueta, decimal AplicadoR3, decimal AnticiposR3,
    decimal AplicadoR4, decimal AnticiposR4, decimal TotalFuente);
```

- `WorkbookLeafInputs.AntExtRev` (`Remuneracion.Core/Models/WorkbookLeafInputs.cs:107` +): `AntExtRevAseInputs?` nullable; `null` = comportamiento anterior (compat por construcción).
- Catálogo de fuentes (patrón `SinonimosEmpresaR1`, `Remuneracion.Core/Models/SinonimosEmpresaR1.cs:26-46`): tabla citada `{EtiquetaBloque, EtiquetasFuente, Cita}` con matcher puro que reutiliza `SonMismaEmpresa` (cero ramas por empresa/período, D-H). Filas iniciales: {BANCO DE OCCIDENTE ← OCCIDENTE, identidad}, {EAAB RECIPROCIDAD ← NUEVO ESQUEMA, cita triple del Plan 34}, {ENEL ← ENEL, identidad}. Sin fila para CIUDAD LIMPIA/ENERBIT (HUE-1).

### §2.3 Contratos (decisión D-02)

1. `IWorkbookLeafInputReader.LeerAntExtRev(Ase ase, Periodo periodo, string rutaRecaudosReversados) → AntExtRevAseInputs` (aditivo; firma orientativa; el H1 fija el nombre exacto). Implementación en `partial` nuevo `ExcelDataReaderWorkbookLeafInputReader.AntExtRev.cs` (mismo tipo, archivo nuevo; cero toque a los 100+ sitios de composición del §1.5 V-03).
2. `ILocalizadorArchivosAse.BuscarRecaudosReversados(string carpetaAse) → string?` (prefijo normalizado `recaudosreversados`: minúsculas + sin diacríticos, misma convención que `BuscarSaldosNotas`/`BuscarRetribucionNegativa` en `ILocalizadorArchivosAse.cs:64-82`). Implementación en `ArchivoFuenteLocator` (mismo `TopDirectoryOnly`, mismo case-insensitive).
3. `ReporteInsumoAse.RecaudosReversados` (nuevo miembro; `ReporteInsumoAse.cs:8-30` hoy R1..RetribucionNegativa).
4. Escritura: `IWorkbookLeafWriter.GenerarWorkbook` intacta en firma (los 100 valores viajan dentro de `WorkbookLeafInputs.AntExtRev`; el writer los consume en la misma pasada atómica, §2.8).

### §2.4 Lector paso a paso (decisión D-04/D-06/D-10)

1. Abrir `Sheet1` por nombre (doctrina Plan 27 D-A; fail-fast que nombra archivo+hoja si ausente).
2. Localizar marcadores por texto normalizado (mayúsculas, sin diacríticos, trim): `Pagos Pendientes por Aplicar` y `Resumen Recaudo Aplicado Por Servicio`. Faltan o invertidos (bloque 3 antes que bloque 2) ⇒ `ERR-FORMATO-FUENTE` (ASE + archivo + marcador).
3. En el bloque 2, recorrer filas: concepto (C normalizado ∈ catálogo §1.2; **C y etiquetas por ENCABEZADO, nunca por índice** — precedente Plan 61 D-H1). En cada fila de concepto/canal/fuente **exigir igualdad D=E=F; si divergen > 0,5 ⇒ `ERR-FORMATO-FUENTE`** (ASE + fuente + concepto + D/E/F observados); **leer F como valor del concepto**. PROHIBIDO mapa quemado D/E/F "por índice": la igualdad es la invariante, no el índice (M4, §0.6).
4. Agrupar por fuente: filas de concepto → acumulador de 4 celdas; total de canal (B=canal, C=Total — informativo, no destino); total de FUENTE (A=fuente, B=Total) cierra el grupo. Σ 4 conceptos ≠ total-fuente > 0,5 ⇒ `ERR-FORMATO-FUENTE` (ASE + fuente + Σ vs total).
5. Fuente del catálogo ausente (cero filas) ⇒ omisión + traza `ANTEXTREV:{fuente}` (1 `Warning` por ASE+fuente, formato Plan 51; propaga a `ResultadoProcesoPeriodo.FuentesAusentes` como `ASE n · ANTEXTREV:{fuente}`). Concepto ausente dentro de fuente presente ⇒ 0 explícito en su celda.
6. Etiqueta de fuente observada no tabulada en el catálogo (+ hay ausencia en el mismo archivo) ⇒ fail-fast A3-gemelo (ASE + `RecaudosReversados` + etiqueta desconocida + fuentes ausentes); sin ausencia ⇒ solo evidencia en `EtiquetasNoReconocidas` (gemelo Plan 50/51). Etiqueta duplicada ⇒ `ERR-FORMATO-FUENTE`. Cero fuentes reconocidas ⇒ estructura rota ⇒ fail-fast.
7. Bloque 3: se lee SOLO como control (RECAUDO OCCIDENTE / RECAUDO OTROS / Total por encabezado; columnas variables). Su divergencia contra Σ-bloque-2 NO bloquea (se registra como hallazgo auditable con ASE + Δ; el H0 fija el umbral con bytes; por defecto ±0,5 informativo, no fail-fast — el bloque 3 no discrimina OCCIDENTE en ASE2 Sept-1, así que no puede juzgar al bloque 2).

### §2.5 Resolvedor de geometría (decisión D-05)

- `ResolvedorDestinoAntExtRev` (Infrastructure, lectura OpenXML/BCL sobre la plantilla que se procesa): dado (etiqueta de bloque B, ASE id 1..5, fila {APLICADO, ANTICIPOS}, reporte {3, 4}) devuelve la celda destino (col C..L por ASE, fila resuelta por etiqueta). Localización: etiqueta de bloque por `B=<etiqueta>`; fila ASE por rótulo ASE1..ASE5 + TOTAL(M); sub-filas por REPORTE 3/4 + APLICADO/ANTICIPOS/TOTAL. **Cero literales de fila/celda en producción** (las etiquetas son DATOS leídos; la T0 es evidencia del fixture). Rejilla de valores C..L; encabezados de reporte C..N (M2).
- Validación estructural pre-escritura (plan-then-write: **falla antes de escribir**): las 5 etiquetas existen y son únicas; cada bloque trae su rejilla completa (10 cols × filas APLICADO/ANTICIPOS + TOTAL con `<f>` SUM por columna); el bloque superior trae sus `<f>` (C3..N3/C4..N4/fila 5/fila 6/P3/P4 según T0). Tercera forma ⇒ `ERR-PLANTILLA` (hoja + identidad esperada + celda encontrada).
- `VerificadorDestinoProtegido` reutilizado (§2-D2 del Plan 59): las 100 celdas de valor son **numéricas sin `<f>`** (M1) y NO son destino con fórmula; el bloque superior + TOTAL por columna SÍ son `<f>` (nunca destino; 130 con `<f>` medidas, sin `<sheetProtection>` — M3). Guard anti-fórmula (`EscribirValorNumerico` → `ERR-PLANTILLA`) intacto y cruzado por las 100 escrituras.

### §2.6 Escritor values-only (decisión D-04/D-08)

- `OpenXmlPlantillaWriter` consume `leaf.AntExtRev` (null ⇒ no-op, comportamiento anterior bit-idéntico): por fuente presente escribe sus 4 valores (2 filas × 2 reportes del bloque de su etiqueta; columna del ASE por etiqueta, nunca por índice). Fuente ausente ⇒ no se toca (cero del template + traza). Concepto ausente ⇒ 0 explícito **numérico** (M1; el template trae numéricos, no texto).
- Values-only estricto: diff `<f>` = ∅ (assert H2). Estilo de destino preservado. Montos tal cual (signo de origen; sin redondeo adicional — el H0 fija si el manual redondea; por defecto sin redondeo, read-back ±0,5). Al barrer `<f>` expandir fórmulas compartidas (`O26` es seguidora `si=3`, M7).
- Values-only estricto: diff `<f>` antes/después = ∅ (assert H2). Estilo de destino preservado. Montos tal cual (signo de origen; sin redondeo adicional — el H0 fija si el manual redondea; por defecto sin redondeo, read-back ±0,5).
- Read-back post-escritura: las 100 celdas ≡ dominio ±0,5; divergencia ⇒ fail-fast (hoja + celda + esperado + escrito). Fórmulas protegidas (bloque superior + TOTAL) byte-idénticas (texto-`<f>`).

### §2.7 Preflight (decisión D-07)

- `ValidadorInsumosPeriodo.Validar` (`Remuneracion.Core/Services/ValidadorInsumosPeriodo.cs:38-157`): por ASE, `BuscarRecaudosReversados` vía `ILocalizadorArchivosAse` (mismo patrón que R1/R2/R4/banco/balance `:67-116`); `null` ⇒ `ReporteAse(aseId, carpetaAse, ReporteInsumoAse.RecaudosReversados)`; resuelto ⇒ firma-chequeo Plan 27 (`AgregarSiFirmaInvalida`, `:165-187`). **En AMBAS quincenas** (sin rama `esQuincena2`; el insumo aplica en Q1 y Q2 por el encargo).
- `FormateadorInsumosFaltantes` (`Remuneracion.Core/Services/FormateadorInsumosFaltantes.cs:8`): nombre administrativo del reporte (`Recaudos reversados por aplicar (ANT EXT-REV)`) + DÓNDE (carpeta ASE + prefijo esperado `RecaudosReversados_`) + QUÉ HACER; jerga prohibida (doctrina Plan 26 D-E).
- Actualización en el mismo cambio: `ReporteInsumoAse.cs:4-7`, `ValidadorInsumosPeriodo.cs:16`, `ILocalizadorArchivosAse` (sin finder) y todo comentario "§V9 / fuera" (grep de cierre los exige; un comentario que bendice la ausencia del insumo es parte del defecto).
- Single-ASE intacto (Plan 26 §V7): `ProcesadorRemuneracion` NO exige este insumo.

### §2.8 Orden dentro de la pasada atómica (decisión)

En `ProcesadorPeriodo.Ejecutar` (preflight `:140-155` → loop ASE `:162+` → `GenerarWorkbook` `:402` → gates `:442+`, ver §1.5 V-06):

1. Preflight (duro, antes de leer — incluye el nuevo insumo).
2. Por ASE: localizar R1/R2/R4 + `RecaudosReversados` (finders; `?? throw` defensa en profundidad) → leer agregados → `LeerLeafInputs` → `LeerAntExtRev` (nuevo; fail-fast `ERR-FORMATO-FUENTE` antes de calcular) → resto de lecturas vigentes.
3. `GenerarWorkbook` a `rutaSinValidar` (política Plan 61 §2-D-H3-ii: escribir→validar): dentro, el pegado ANT EXT-REV corre con el resto de escrituras values-only, ANTES de las lecturas-oráculo/gates.
4. Gates de datos (novedad, Plan 61) + 2.7 + coherencia ANT EXT-REV (§2.9): clasifican, no bloquean la entrega.

### §2.9 Gates, observabilidad y CodigoError

- **Gates nuevos (Core, aditivos):** coherencia ANT EXT-REV por ASE (Σ 4 conceptos por fuente ≡ total-fuente ±0,5; Σ bloques leídos ≡ lo escrito ±0,5) en `ValidadorBasico` junto a los juicios vigentes (SRP: no se toca el juicio R1/R2/R4); `ClasificarDivergenciasDeclaradas` lo hereda (divergencia ANT EXT-REV = NO declarada ⇒ falla en duro, Plan 42 D-F).
- **Observabilidad (HU-14):** 1 `Warning` por (ASE, fuente ausente) + hito `Information` por ASE (`ANT EXT-REV leído: n fuentes presentes, m ausentes`); `RunId` correlaciona (ámbito vigente `ProcesadorPeriodo.cs:100-104`).
- **CodigoError** (`Remuneracion.Core/Errors/CodigoError.cs:13-38`): estructura rota del insumo ⇒ `FormatoFuente = "ERR-FORMATO-FUENTE"`; plantilla sin destino ⇒ `Plantilla = "ERR-PLANTILLA"` / `SinSlot = "ERR-SIN-SLOT"`; preflight ⇒ `FuenteNoEncontrada = "ERR-FUENTE-NO-ENCONTRADA"` (salida 2, Plan 26 D-B). Sin código nuevo.

### Architecture Validation Certificate (diseño)

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Reader = evidencia (lee por encabezado/marcador, omite con huella, reporta); Resolvedor = etiquetas→celda; Writer = values-only; Validador = juicio; Preflight = existencia+firma. Cada pieza una sola razón de cambio |
| **OCP** | ✅ Fuente nueva con cita = fila en el catálogo, cero código; geometría futura con la misma gramática (B-etiqueta + ASE + REPORTE + APLICADO/ANTICIPOS) se absorbe sin código nuevo |
| **LSP/DIP** | ✅ Contratos `IWorkbookLeafInputReader`/`ILocalizadorArchivosAse`/`IPlantillaWriter`/`IValidador` intactos en firma salvo adiciones; `WorkbookLeafInputs.AntExtRev` nullable (defaults = anterior); cambios detrás de las mismas abstracciones |
| **Best practices** | ✅ Contract-first (H0 congela bytes antes del juicio); fail-fast accionable (ASE + reporte + marcador/etiqueta + esperado vs obtenido); tolerancia ±0,5 solo en comparación; values-only; D-H (cero literales/ramas por período/empresa/fuente en producción nueva); CRLF; sin commits; prohibido `<v>` stale |
| **Performance** | ✅ H0 acotado BCL + lectura solo-copias; lectura O(filas del bloque 2) por ASE una vez por corrida; resolución de geometría O(bloques × filas) una vez; cero workbooks extra en runtime; sin Excel en los gates |

---

## §3. SPEC (requisitos numerados + criterios de aceptación)

### S0 — Gates transversales (toda H)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa trx + **set-diff de rojos vs base de entrada**: altas/bajas una por una con causa citada; **rojo nuevo fuera de lista ⇒ STOP**. 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados, 0 tolerancias ampliadas; re-anclajes solo con evidencia.
- G-HUELLA: huella v4 `a7476ee2…8c55` intacta antes/después de cada H (el SHA de archivo es informativo).
- G-FORMULA: `<f>` byte-idéntica en todo el diff salvo allow-list declarada (cero; este plan no toca `<f>`): cero `new CellFormula`, cero `InsertRow`/`DeleteRow`. Guard anti-fórmula intacto.
- G-QUEMADO: cero literales de fila/celda/letra y cero ramas por ASE/período/quincena/empresa/fuente en los caminos tocados —incluida la forma escondida en mapas y mensajes— (grep de cierre + candado §4-H5).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale) o la fuente ya leída; **cero workbooks extra en runtime**; **sin Excel en el gate**.
- G-TRAZA: cada novedad y cada error nuevo nombra reporte + hoja + ASE + concepto/celda (+ observado vs esperado + Δ donde aplique). Prohibida jerga en lo administrativo.
- G-PROC: cero commits; cero `git checkout` / `git reset`; CRLF sin mezclar en lo tocado; Excel cerrado durante corridas y T0.
- G-H0-READONLY: H0 no contiene código de producción (solo fixture `*T0Plan64*` + `Docs/T0-Plan64-Ant-Ext-Rev.md`).

### REQ — Requisitos (vinculantes; todo número viaja con su derivación fuente × etiqueta)

| ID | Requisito | Grupo | Criterio de aceptación |
|---|---|---|---|
| REQ-01 | Geometría v4 congelada por bytes (bloque superior, 5 etiquetas, rejilla 100 celdas, naturaleza por celda) | H0 | T0.1–T0.2 verdes; §1.1 re-medido o enmendado con STOP previo |
| REQ-02 | Estructura Sheet1 congelada (marcadores, bloque 2 con C/etiquetas por encabezado + invariante D=E=F, totales de canal/fuente, bloque 3 como control con 8 archivos sin OCCIDENTE — lista M5) + universo 30/30 | H0 | T0.3–T0.4 verdes; HUE-3 CERRADO por M4 |
| REQ-03 | Paridad numérica §1.4 re-derivada + invariante Σ-conceptos == total-fuente ±0,5 | H0 | T0.5 verde; mismatch ⇒ STOP |
| REQ-04 | `LeerAntExtRev` en `partial` nuevo: marcadores, conceptos por encabezado (C + etiquetas), **igualdad D=E=F con fail-fast > 0,5 y lectura de F (M4; prohibido mapa quemado por índice)**, agrupación por fuente, signo tal cual, ausente = traza, concepto ausente = 0, desconocida/duplicada/cero-fuentes/Σ ⇒ fail-fast | H1 | Lectura pura de los 30 archivos ≡ T0 ±0,5; 0 warnings |
| REQ-05 | Catálogo citado + `AntExtRevAseInputs` + `WorkbookLeafInputs.AntExtRev` nullable; HUE-1 (CIUDAD LIMPIA/ENERBIT sin fila; si vienen ⇒ fail-fast) | H1 | Tests de catálogo (identidad + sinonimia + no-tabulado-se-iguala-a-sí-mismo + injerto sin cita ⇒ rojo) |
| REQ-06 | Resolvedor por etiquetas + validación estructural plan-then-write (`ERR-PLANTILLA`); prohibido C11..L12 en producción | H1 | T0 como evidencia, no como contrato; injerto con literal ⇒ rojo |
| REQ-07 | Escritor values-only 100 celdas en la pasada atómica (§2.8); diff `<f>` = ∅; read-back ±0,5; protegidas intactas | H2 | Corridas 2026091 + 2026101: ANT EXT-REV ≡ manual por bloque/ASE ±0,5 |
| REQ-08 | Preflight en ambas quincenas + firma + mensaje administrativo + reescritura §V9 y comentarios | H3 | Período sin el archivo ⇒ UN `ERR-FUENTE-NO-ENCONTRADA` agregado, antes de leer, sin salida |
| REQ-09 | Comparador sin exclusión ANT EXT-REV + ítem 8 actualizado + goldens por bloque/ASE contra los 4 manuales | H4 | Suite del comparador verde con la hoja comparada; test negativo anti-exclusión |
| REQ-10 | Candado anti-quemado (estático + injerto + conductual) + sonda de etiquetas + doctrina de fixtures | H5 | Candados verdes; review APROBADO |

### CA — Criterios de aceptación (vinculantes)

- **CA-01:** H0 verde (6 frentes, 18/18 + `Docs/T0-Plan64-Ant-Ext-Rev.md`) + ENMIENDA M1–M6 aplicada + APROBADO del Ingeniero. H1 habilitado bajo esta enmienda.
- **CA-02:** los 30 insumos leídos en puro ≡ T0 ±0,5 (D/E/F, Σ-fuente, totales de canal informativos).
- **CA-03:** corridas multi-ASE 2026091 (Sept-1) y 2026101 (Oct-1) en AMBAS quincenas donde aplique: ANT EXT-REV ≡ manual por bloque/ASE ±0,5; `ANTICIPOS USUARIOS` y `L125:L131` (M6: referencian L125:L129, L130 vacía, L131 consumidora) recalculan solos (sin escritura, sin exclusión nueva).
- **CA-04:** período sin `RecaudosReversados_*` ⇒ fail-fast agregado pre-lectura con mensaje administrativo (QUÉ/DÓNDE/QUÉ HACER, sin jerga).
- **CA-05:** comparador compara ANT EXT-REV (exclusión retirada, ítem 8 actualizado); 4 manuales como goldens por bloque/ASE.
- **CA-06:** build 0 warnings; suite set-diff exacto; huella v4 intacta; review independiente **APROBADO**.

---

## §4. TASKS (fases H0..H5 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero código de producción)

Fixture: `Remuneracion.IntegrationTests/AntExtRevT0Plan64Tests.cs` (BCL ZIP+XML + ExcelDataReader solo-lectura sobre copias; rutas explícitas; prohibido `Insumos.Raiz()` en lo nuevo). Evidencia: `Docs/T0-Plan64-Ant-Ext-Rev.md`.

| ID | Tarea | Verifica |
|---|---|---|
| T0.1 | Geometría v4 por bytes: bloque superior (texto-`<f>` C3..N3/C4..N4/fila 5/fila 6/P3/P4), 5 etiquetas B8/B16/B24/B32/B40, rejilla C..L × filas de valor (reportes C..N — M2), naturaleza numérica de las 100 celdas (M1), sin `<sheetProtection>` (M3) | REQ-01 |
| T0.2 | 100 %-fórmulas de ANTICIPOS USUARIOS (barrido `<f>` + referencias D4/M4/D20/M20…; `O26` seguidora `si=3` — M7) + `CONSOLIDADO!L125:L129` como referenciadoras, L130 vacía, L131 consumidora (M6) | REQ-01 |
| T0.3 | Sheet1 por archivo (30/30): marcadores (fila exacta como evidencia), invariante D=E=F sin encabezados semánticos (M4; C/etiquetas por encabezado), filas de concepto/canal/fuente, bloque 3 (columnas variables; 8 archivos sin OCCIDENTE — lista M5) | REQ-02 |
| T0.4 | Universo de fuentes 30/30 (ENEL/NUEVO ESQUEMA/OCCIDENTE) + ausencia 30/30 de CIUDAD LIMPIA/ENERBIT + 0 en los 4 manuales | REQ-02 + HUE-1 |
| T0.5 | Paridad §1.4 re-derivada `fuente × etiqueta` (Sept-1 ASE1/ASE4, Julio-2 ASE4, Octubre-1 ASE4) + invariante Σ == total-fuente ±0,5 | REQ-03 |
| T0.6 | Divergencia docx-vs-encargo (D-10): cita exacta del docx ("Resumen") + veredicto bloque-2-fuente/bloque-3-control con bytes | REQ-03 |
| T0.7 | Lista cerrada de Δ de suite + huella v4 antes/después de H0 | S0 |

- **Puerta H0→H1:** T0.1–T0.7 verdes + S1.5-equivalente sin STOP + APROBADO del Ingeniero (incluidas D-01..D-10 del §0.5).

### H1/F1 — Lector + modelos + catálogo + resolvedor (traza a §2.2–§2.5)

| ID | Tarea | Archivo | Traza |
|---|---|---|---|
| H1.1 | Core: `AntExtRevAseInputs` + `AntExtRevFuenteInputs` (puros) + catálogo de fuentes con etiqueta de bloque y citas + matcher por `SonMismaEmpresa` | `Remuneracion.Core/Models/AntExtRevAseInputs.cs`, `.../CatalogoFuentesAntExtRev.cs` | REQ-05, D-03 |
| H1.2 | Contrato aditivo `LeerAntExtRev` en `IWorkbookLeafInputReader` + `WorkbookLeafInputs.AntExtRev` nullable | `Remuneracion.Core/Interfaces/IWorkbookLeafInputReader.cs`, `Remuneracion.Core/Models/WorkbookLeafInputs.cs` | REQ-05, D-02 |
| H1.3 | `partial` nuevo del reader: §2.4 paso a paso (marcadores, encabezados, agrupación, Σ-fuente, traza, A3-gemelo, fail-fast D-06) | `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.AntExtRev.cs` (nuevo) | REQ-04, D-04/D-06/D-10 |
| H1.4 | Resolvedor por etiquetas + validación estructural plan-then-write (sin literales C11..L12) | `Remuneracion.Infrastructure/Excel/ResolvedorDestinoAntExtRev.cs` (nuevo) | REQ-06, D-05 |
| H1.5 | Tests focalizados (rojo→verde honesto): 30 archivos ≡ T0; marcadores invertidos ⇒ fail; desconocida+ausente ⇒ fail; solo-evidencia sin ausencia; duplicada ⇒ fail; **D≠E≠F>0,5 ⇒ fail (M4; se lee F)**; Σ≠total ⇒ fail; sin columna → 0 (regresión M-B); **columna movida ⇒ mismo valor solo si D=E=F se preserva (la invariante, no el índice)** | `Remuneracion.IntegrationTests/AntExtRevLecturaH1Plan64Tests.cs` (nuevo) | REQ-04/05/06 |
| H1.6 | Greps G-QUEMADO (parcial H1) + huella + 0 warnings | — | S0 |

- **Puerta H1→H2:** REQ-04/05/06 verdes + CA-02 + 0w + huella + greps.

### H2/F2 — Escritor values-only + read-back (traza a §2.6/§2.8)

| ID | Tarea | Archivo | Traza |
|---|---|---|---|
| H2.1 | Consumo de `leaf.AntExtRev` en la pasada atómica (§2.8): presentes → 4 valores; ausentes → intacto + traza; concepto ausente → 0 | `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` (punto citado en H1 contra `ProcesadorPeriodo.Ejecutar`) | REQ-07, D-04 |
| H2.2 | Read-back ±0,5 de las 100 celdas + protegidas byte-idénticas (texto-`<f>`) + diff `<f>` = ∅ | mismo writer + `Remuneracion.IntegrationTests/AntExtRevEscrituraH2Plan64Tests.cs` (nuevo) | REQ-07, D-08 |
| H2.3 | e2e 2026091 + 2026101: ANT EXT-REV ≡ manual por bloque/ASE ±0,5 (**vacío ≡ 0** solo en `2026092!J20`-análogos declarados; §0.6); ANTICIPOS USUARIOS y L125:L131 recalculan (sin escritura) | corridas CLI a temp fresca | CA-03 |
| H2.4 | Greps G-QUEMADO (total H2) + huella + 0 warnings | — | S0 |

- **Puerta H2→H3:** REQ-07 verde + CA-03 + 0w + huella + diff `<f>` = ∅.

### H3/F3 — Preflight + firma + doctrina §V9 (traza a §2.7)

| ID | Tarea | Archivo | Traza |
|---|---|---|---|
| H3.1 | `BuscarRecaudosReversados` (prefijo normalizado `recaudosreversados`) | `Remuneracion.Core/Interfaces/ILocalizadorArchivosAse.cs`, `Remuneracion.Infrastructure/FileSystem/ArchivoFuenteLocator.cs` | REQ-08, D-07 |
| H3.2 | `ReporteInsumoAse.RecaudosReversados` + nombre administrativo en `FormateadorInsumosFaltantes` + exigencia en `ValidadorInsumosPeriodo` (AMBAS quincenas, sin rama) + firma-chequeo | `Remuneracion.Core/Models/ReporteInsumoAse.cs`, `Remuneracion.Core/Services/FormateadorInsumosFaltantes.cs`, `Remuneracion.Core/Services/ValidadorInsumosPeriodo.cs` | REQ-08, D-07 |
| H3.3 | Reescritura de doctrina §V9 y comentarios mentirosos (grep de cierre los exige) + checklist paridad preflight↔runtime | `.opencode/project-context.md` (H5-docs lo cierra; H3 deja el diff citado) + comentarios en código | REQ-08, R-10 |
| H3.4 | Tests: período sin el archivo ⇒ UN `ERR-FUENTE-NO-ENCONTRADA` agregado pre-lectura; firma inválida ⇒ variante administrativa; Q1 también lo exige | `Remuneracion.IntegrationTests/AntExtRevPreflightH3Plan64Tests.cs` (nuevo) | CA-04 |

- **Puerta H3→H4:** REQ-08 verde + CA-04 + 0w + mensaje sin jerga (grep).

### H4/F4 — Comparador + BrechaDeclarada + goldens (traza a D-09)

| ID | Tarea | Archivo | Traza |
|---|---|---|---|
| H4.1 | Retirar la exclusión ANT EXT-REV del comparador (la de ANTICIPOS USUARIOS queda: sigue sin escribirse) | `Remuneracion.Infrastructure/Excel/ComparadorSalidaVsManual.cs` (si la exclusión vive allí) — hoy la exclusión medida vive en el TEST `:159-162`; el H4 fija dónde vive la forma vigente y la retira solo para ANT EXT-REV | REQ-09 |
| H4.2 | Actualizar `BrechaDeclarada` ítem 8: ANT EXT-REV sale del non-goal; ANTICIPOS USUARIOS permanece con su motivo | `Remuneracion.IntegrationTests/ComparadorSalidaVsManualTests.cs:158-162` | REQ-09 |
| H4.3 | Goldens por bloque/ASE contra los 4 manuales (Julio-2, Septiembre-1, Septiembre-2, Octubre-1) ±0,5 / f-vs-f (**vacío ≡ 0** en `2026092!J20`; §0.6) | `Remuneracion.IntegrationTests/AntExtRevComparadorH4Plan64Tests.cs` (nuevo) | CA-05 |
| H4.4 | Test negativo: divergencia ANT EXT-REV sin exclusión ⇒ rojo (anti-falso-verde R-07) | mismo fixture | R-07 |

- **Puerta H4→H5:** REQ-09 verde + CA-05 + set-diff sin rojos fuera de lista.

### H5/F5 — Prevención sistémica + cierre (traza a REQ-10)

| ID | Tarea | Traza |
|---|---|---|
| H5.1 | Candado anti-quemado extendido (estático: literal de celda/fila/ventana/letra + rama ASE/período/quincena/empresa/fuente en los caminos tocados; injertos en rojo por forma; conductual: mismo insumo reordenado coloca idéntico + geometría desplazada resuelve) | G-QUEMADO, R-01 |
| H5.2 | Sonda de etiquetas del insumo (encabezado no mapeado ⇒ hallazgo `ASE n · RecaudosReversados · Sheet1 · {etiqueta}`; coexistencia con ausencia ⇒ fail-fast, precedente Plan 50 A3) | R-04 |
| H5.3 | Doctrina de fixtures (prohibido inventar valores del insumo; asserts contra T0 medido; prohibidos asserts tautológicos, precedente Plan 61 H5.4) + candado textual | S0 |
| H5.4 | Suite trx + set-diff final; code-review independiente (`code-reviewer`, este plan como spec): **APROBADO** requerido | CA-06 |
| H5.5 | Doctrina en `.opencode/project-context.md` (solo lo probado: lectura por encabezado/marcador, catálogo, opcionalidad con traza, preflight en ambas quincenas, comparador sin exclusión, HUE-1) + cierre del plan con desviaciones fechadas; cero código de producción en H5-docs | CA-06 |

**Orden sugerido:** H0 → H1 → H2 → H3 → H4 → H5. Regresión continua (suite + huella en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero). Excel cerrado durante corridas y T0.

---

## §5. Gates de cierre (medibles — todos o no cierra)

- G1: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**; suite verde con set-diff exacto vs base H0 (rojo fuera de lista = STOP).
- G2: corridas CLI a temp fresca (rutas explícitas, Excel cerrado):
  - `dotnet run --project Remuneracion.Cli -- --periodo 2026091 --carpeta "Docs/Prueba Septiembre 1/Insumos" --plantilla "Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx" --salida <temp>\2026091-ant-ext-rev.xlsx --cinco-ase`
  - `dotnet run --project Remuneracion.Cli -- --periodo 2026101 --carpeta "Docs/Prueba Octubre-1/Insumos" --plantilla <v4> --salida <temp>\2026101-ant-ext-rev.xlsx --cinco-ase`
  - (Flags exactos según `EjecutorCli`; el H2 fija la línea real contra `--help` sin cambiar la semántica.)
  - Criterio: ANT EXT-REV ≡ manual por bloque/ASE ±0,5 (vacío ≡ 0 solo en celda declarada §0.6); read-back verde; protegidas byte-idénticas; `ANTICIPOS USUARIOS` y `L125:L131` (M6) recalculan sin escritura.
- G3: período sin `RecaudosReversados_*` ⇒ `ERR-FUENTE-NO-ENCONTRADA` agregado pre-lectura (salida 2, línea `RESULTADO ERROR` grepable), sin workbook abierto ni salida creada.
- G4: comparador sin exclusión ANT EXT-REV: 0 inesperadas en la hoja contra los 4 manuales (ítem 8 actualizado; ANTICIPOS USUARIOS permanece declarado).
- G5: candados H5 verdes (incluidos injertos en rojo); huella v4 intacta; review independiente **APROBADO**.

Artefactos: `Docs/T0-Plan64-Ant-Ext-Rev.md` + fixtures `AntExtRevT0Plan64Tests` / `AntExtRevLecturaH1Plan64Tests` / `AntExtRevEscrituraH2Plan64Tests` / `AntExtRevPreflightH3Plan64Tests` / `AntExtRevComparadorH4Plan64Tests` + trx `p64-h*-suite.trx` + setdiffs + salidas `<temp>` + reporte de cierre en este §8.

---

## §6. Impacto en tests existentes y plan de actualización

| Frente | Impacto | Tratamiento (sin debilitar: 0 Skip, 0 borrados, 0 tautologías) |
|---|---|---|
| Fixtures sin el archivo `RecaudosReversados_*` | Composiciones de `ProcesadorPeriodo` en tests (100+ sitios, §1.5 V-03) que resuelven períodos de prueba sin el insumo ⇒ preflight rojo nuevo | H3 inventaría TODOS los sitios que componen períodos sin el insumo: o se les provee el insumo real (copia temporal, nunca fixture en repo) o se les declara el faltante esperado con causa citada. Rojo fuera de esa lista = STOP |
| Preflight (`PreflightInsumosTests`, `BasePeriodoPreflightTests`) | La superficie crece (+1 insumo × 5 ASE, ambas quincenas); tests "lista vacía" con períodos completos siguen verdes solo si esos períodos traen el archivo | H3 re-ancla con evidencia (el período completo trae el archivo ⇒ verde; si no lo trae ⇒ el test pasa a exigir el faltante). Prohibido nombre exacto quemado (por prefijo normalizado, precedente Plan 59 H1.6) |
| Comparador (`ComparadorSalidaVsManualTests.BrechaDeclarada:158-162`) | Ítem 8 pierde ANT EXT-REV; las corridas históricas sin la escritura enrojecen en la hoja | H4 actualiza el ítem con cita T0; los goldens nuevos fijan el verde. ANTICIPOS USUARIOS intacto |
| Goldens e2e (julio/agosto/septiembre/octubre) | Las salidas ganan 100 valores donde había ceros ⇒ deltas DECLARADOS en ANT EXT-REV (+ arrastre recalculado en ANTICIPOS USUARIOS/L125:L131 — M6 —, que NO son deltas de escritura sino recálculo) | H2/H4 publican la tabla corrida→nueva justificada por celda (derivación `fuente × etiqueta`); fuera de esa tabla = STOP |
| Comentarios §V9 y "fuera" | Quedan falsos (`ValidadorInsumosPeriodo.cs:16`, `ReporteInsumoAse.cs:4-7`, `ILocalizadorArchivosAse` sin finder) | H3 los reescribe en el mismo cambio (grep de cierre) |
| `AnclasDerivadasH2` / `SalidaSinValidarO4` / `EntregaConNovedadesH3` | Si algún test asume "ANT EXT-REV intacta en ceros" o ausencia de salida ante divergencia, cambia de premisa | H2/H4 los inventarían (T-inventario en H0 si el grep los encuentra) y los re-anclan preservando intención (no-certificar ≠ no-escribir, Plan 61 H3.3) |

---

## §7. Riesgos y rollback

Ver §0.4 (R-01..R-10). Gates por fase:

- **Gate H0→H1:** T0.1–T0.7 verdes (18/18) + ENMIENDA M1–M6 aplicada + APROBADO del Ingeniero. **Puerta ABIERTA.** Baseline ratificado: base 1799 (1776 + 23 rojos de entrada, T0 §7.3) → H0 1817 (+18, 0 nuevos).
- **Gate H1→H2:** REQ-04/05/06 + CA-02 + 0w + huella + greps G-QUEMADO.
- **Gate H2→H3:** REQ-07 + CA-03 + 0w + huella + diff `<f>` = ∅.
- **Gate H3→H4:** REQ-08 + CA-04 + mensaje sin jerga + checklist de paridad.
- **Gate H4→H5:** REQ-09 + CA-05 + set-diff sin rojos fuera de lista.
- **Gate H5/cierre:** REQ-10 + CA-06 + review APROBADO + doctrina solo-probada.

**Rollback:** cada H es reversible por separado (modelos nullable + preflight aditivo + exclusión versionada). Si H2 enrojece goldens fuera de tabla ⇒ se revierte el consumo del writer (el `null` restaura el comportamiento anterior bit-idéntico) sin tocar el lector; si H3 enrojece fixtures fuera de lista ⇒ se revierte la exigencia del preflight (el finder permanece, sin exigir) hasta decisión. Nunca parche por ASE/fuente/período.

### §7.8 Corrida final H5 (suite completa + set-diff)

Fecha: 2026-10-30. Excel cerrado; una sola corrida (sin procesos concurrentes).

```
dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx
  -> Build succeeded. 0 Warning(s) / 0 Error(s)

dotnet test Remuneracion.IntegrationTests/Remuneracion.IntegrationTests.csproj --no-build ^
  --logger "trx;LogFileName=p64-h5-suite.trx" ^
  --results-directory "Remuneracion.IntegrationTests/TestResults"
  -> Failed: 23, Passed: 1845, Skipped: 0, Total: 1868 (21 m 11 s)
  -> Remuneracion.IntegrationTests/TestResults/p64-h5-suite.trx
```

- **Baseline:** `Docs/T0-Plan64-Ant-Ext-Rev.md` §7.3 (23 rojos de entrada, ajenos a este plan).
- **Set-diff de rojos:** **0 nuevos / 0 desaparecidos** (el conjunto es idéntico al baseline). **0 `Skip`**. Detalle: `Remuneracion.IntegrationTests/TestResults/p64-h5-setdiff.txt`.
- **Fixtures del plan:** **69 casos, 0 fallos** (`AntExtRevT0Plan64Tests` 18 · `AntExtRevLecturaH1Plan64Tests` 20 · `AntExtRevEscrituraH2Plan64Tests` 6 · `AntExtRevPreflightH3Plan64Tests` 5 · `AntExtRevComparadorH4Plan64Tests` 6 · `AntExtRevCandadoPrevencionH5Plan64Tests` 14), 0 `Skip`.
- **Huella v4:** intacta — `AntExtRevT0Plan64Tests.T0_07_HuellaV4_Intacta` y `HuellaPlantillaVigenteV4Plan59Tests` (incluida `V4_5_INYECCION…`) en verde.
- **G-QUEMADO de cierre:** 0 literales de coordenada de la v4 (`C11`/`C12`/`D11`/`D12`/`C19`/`C27`/`C35`/`C43`/`B8`/`B16`/`B24`/`B32`/`B40`) y 0 ramas `NumeroQuincena ==|!=` en la producción nueva; 0 `new CellFormula`/`InsertRow`/`DeleteRow`; 0 comentarios «`RecaudosReversados` FUERA»/«§V9» en código de producción.
- **Desviación operativa declarada (no afecta el veredicto):** una corrida intermedia se ejecutó **en paralelo** con otra suite sobre el mismo `bin`; el log compartido (`remuneracion_log_YYYYMMDD.txt`) generó rollovers `_00N` y añadió 1 rojo **artificial** (`ParidadCliTests.Paridad_5Ase_Q1_CliVsDirecto_ValoresYSnapshotsIguales`, «encontradas 0» líneas con el RunId). En entorno limpio ese fixture pasa **5/5**. La corrida reportada arriba es la limpia; el rojo no era del código ni del plan.
- **Review independiente (`code-reviewer`):** **PENDIENTE** — es el único criterio de CA-06 que no cierra el implementer (la puerta H5 queda REQ-10 + CA-06 **parcial** hasta ese veredicto).


---

## §8. Checklist de cierre y qué NO se hizo

### Checklist (el implementer lo marca con evidencia)

- [x] H0 verde (T0.1–T0.7, 18/18) + `Docs/T0-Plan64-Ant-Ext-Rev.md` + trx base citado (1799: 1776 + 23 rojos de entrada T0 §7.3) + H0 1817 (+18, 0 nuevos) + huella v4 antes/después + ENMIENDA M1–M6 aplicada (§0.6).
- [x] H1 verde (REQ-04/05/06 + CA-02) + 0 warnings + greps G-QUEMADO. Evidencia 2026-10-30: `AntExtRevLecturaH1Plan64Tests` verde (30 insumos ≡ total por fuente de T0 §4.3; paridad celda a celda contra los 20 archivos con manual resolviendo la celda por `ResolvedorDestinoAntExtRev`; negativos de marcadores invertido/ausente, terna que diverge, Σ≠total, etiqueta desconocida con y sin ausencia, fuente duplicada y **columna de valores movida**); build 0 warnings.
- [x] H2 verde (REQ-07 + CA-03: 2026091 y 2026101 por bloque/ASE ±0,5) + diff `<f>` = ∅ + read-back + huella. Evidencia 2026-10-30: `AntExtRevEscrituraH2Plan64Tests` 6/6 (writer + CLI real `--cinco-ase` a temp fresca, exit 0 en ambos períodos); suite 1843 (1820 verdes + los 23 rojos de entrada de §7.3, set-diff exacto); build 0 warnings; huella v4 intacta (`HuellaPlantillaVigenteV4Plan59Tests` 5/5).
- [x] H3 verde (REQ-08 + CA-04: fail-fast agregado, mensaje administrativo, firma, §V9 reescrito) + checklist de paridad. Evidencia 2026-10-30: `AntExtRevPreflightH3Plan64Tests` verde (período sin el archivo ⇒ UN `ERR-FUENTE-NO-ENCONTRADA` agregado antes de leer y sin salida; firma inválida ⇒ variante administrativa; QUÉ/DÓNDE/QUÉ HACER sin jerga; finder con el mismo archivo que el runtime y normalización del nombre). §V9 reescrito en `.opencode/project-context.md` (lo cierra H5.5).
- [x] H4 verde (REQ-09 + CA-05: exclusión retirada, ítem 8 actualizado, goldens 4 manuales, anti-falso-verde). Evidencia 2026-10-30: `AntExtRevComparadorH4Plan64Tests` verde (paridad por bloque/ASE contra los 4 manuales en valores y texto de fórmula; divergencia de ANT EXT-REV reportada SIN exclusión; ítem 8 declara solo `ANTICIPOS USUARIOS`).
- [x] H5 verde (REQ-10: candados + sonda + doctrina de fixtures + set-diff + doctrina solo-probada). Evidencia 2026-10-30: `AntExtRevCandadoPrevencionH5Plan64Tests` **14/14** (estático con allow-list cerrada + injertos en rojo incluido copia temporal revertida + conductual «insumo reordenado coloca idéntico» y «geometría desplazada resuelve» + sonda de etiquetas + candado de doctrina de fixtures); suite **1868** (1845 verdes + los 23 rojos de entrada de §7.3, **set-diff 0 nuevos / 0 desaparecidos**, 0 Skip) → `p64-h5-suite.trx` + `p64-h5-setdiff.txt` (§7.8); build 0 warnings; huella v4 intacta; greps G-QUEMADO de cierre en 0. Review independiente: **PENDIENTE** (CA-06 parcial, ver §8-Desviaciones).
- [x] §2.9 (requisito pendiente tras H5) verde: gates de coherencia ANT EXT-REV por ASE en `ValidadorBasico` (aditivos, junto a los juicios vigentes; SRP: juicio R1/R2/R4 intacto) — Σ de los 4 conceptos por fuente ≡ total-fuente ±0,5 y Σ bloques leídos ≡ lo escrito ±0,5, ambos sobre el dominio YA leído (G-WORKBOOK: cero workbooks en el gate) — + herencia en `ClasificarDivergenciasDeclaradas` (divergencia ANT EXT-REV = NO declarada ⇒ falla en duro, Plan 42 D-F; cero entradas declarables nuevas) + ficha de novedad con la hoja real y los montos comparados (G-TRAZA). Política Plan 61 §2-D-H3-ii: gates de DATOS ⇒ novedad (no retienen el Excel); los de ESTRUCTURA siguen en duro; orden §2.8 respetado (corre después de escribir y antes del oráculo). Evidencia 2026-10-30: `AntExtRevGatesCoherenciaTests` **6/6** (rojo→verde honesto: Σ-fuente diverge ⇒ novedad con ASE + fuente + observado/esperado + Δ; dinero leído sin bloque destino ⇒ ídem por ASE —el caso que el gate por fuente NO ve—; verde del corpus real 30/30; no-op sin lectura del insumo; clasificación NO declarada; ficha con hoja/ASE/montos); familia `AntExtRev*` **75/75**; build 0 warnings; huella v4 intacta (`HuellaPlantillaVigenteV4Plan59Tests` + `T0_07` 6/6); suite **1874** (1851 verdes + los 23 rojos de entrada de §7.3), **set-diff 0 nuevos / 0 desaparecidos**, 0 Skip → `p64-29-suite.trx` + `p64-29-setdiff.txt`.
- [x] Desviaciones con fecha/motivo/evidencia (ver §8-Desviaciones, 2026-10-30).

### §8-Desviaciones declaradas (2026-10-30)

Todas con fecha, motivo y evidencia. Ninguna cambia alcance ni decisiones D-01..D-10.

| # | Desviación | Motivo | Evidencia |
|---|---|---|---|
| DV-H5-1 | **H5.2 añade código de producción**: `AntExtRevAseInputs.HallazgosEtiquetasNoMapeadas` (lista durable) + un `Warning` por etiqueta sin mapear con la forma exigida por H5.2 | El hallazgo `ASE n · RecaudosReversados · Sheet1 · {etiqueta}` **no existía como texto durable**: H1 solo exponía la evidencia `ANTEXTREV:{etiqueta}` y el fail-fast A3. Se añade de forma **aditiva** sin tocar el contrato de H1 (`EtiquetasNoReconocidas` intacto) ni la lectura | `H5_2_Sonda_EtiquetaSinMapear_HallazgoConFormaAdministrativa` + `H5_2_Sonda_Los30InsumosReales_NoTienenEtiquetasSinMapear` + `H5_2_Sonda_EtiquetaSinMapearConAusencia_FailFast` |
| DV-H5-2 | **Se retiró un assert** de `AntExtRevLecturaH1Plan64Tests`: `Assert.Equal(125m, fuente.SumaConceptos)` | Tautología (la propiedad es la suma de los cuatro campos ya medidos en las líneas anteriores); doctrina de fixtures H5.3, precedente Plan 61 H5.4. **No** se borró ningún test ni se amplió ninguna tolerancia: el cierre Σ ≡ total lo hace el lector (fail-fast) y el total de fuente sí se contrasta | `AntExtRevLecturaH1Plan64Tests.Lectura_ConceptoAusenteEnFuentePresente_EsCeroExplicito` (verde) + candado `H5_3_Candado_DoctrinaDeFixtures` |
| DV-H5-3 | **Allow-list cerrada del candado** con 5 entradas motivadas: `F199`/`R1`/`T0` del catálogo citado, `R1`/`R2`/`R4` en los mensajes de progreso de `ProcesadorPeriodo`, y `B` como columna declarada de etiquetas en el resolvedor | Son **citas, nombres administrativos de reportes o gramática declarada**, no coordenadas de destino; sin lista cerrada el candado no puede distinguirlos de un quemado real | `H5_1_CaminosTocados_SinCoordenadasNiRamasQuemadas` + `H5_1_ElAnalizador_NoEnrojeceConLaMaquinariaLegitima` + `H5_1_Injerto_*` |
| DV-H5-4 | El **corpus del candado de doctrina excluye el propio fixture de candado** (5 fixtures H0..H4 + el de H5) | Auto-juicio circular: el fixture de H5 contiene las formas prohibidas como **texto de sus pruebas de injerto**, igual que el candado de período del Plan 59 no se juzga a sí mismo | `H5_3_Candado_DoctrinaDeFixtures` (`Assert.Equal(5, fixtures.Count)`) |
| DV-H5-5 | **Desviación operativa (no del plan):** una corrida intermedia se ejecutó **en paralelo** con otra suite sobre el mismo `bin` y añadió 1 rojo artificial (`ParidadCliTests.Paridad_5Ase_Q1...`) por el log compartido con rollover `_00N` | Error de operación del implementer (dos `dotnet test` a la vez); el test exige ≥2 líneas con el RunId en el log diario y la corrida concurrente partió el archivo. En entorno limpio pasa 5/5. **No se tocó el test ni el código** para “arreglarlo” | `p64-h5-setdiff.txt` + corrida limpia §7.8 (`p64-h5-suite.trx`: 23 rojos = baseline exacto) |

### Qué NO se hizo (vinculante)

1. No se escribió `ANTICIPOS USUARIOS` ni `CONSOLIDADO!L125:L131` (M6: 100 % fórmulas; recalculan solas).
2. No se tocó single-ASE (`ProcesadorRemuneracion` intacto salvo lo que H3 demuestre con archivo:línea, con justificación).
3. No se inventó ningún mapeo (HUE-1 declarado; etiqueta no tabulada que VIENE ⇒ fail-fast, nunca 0 silencioso).
4. No se hardcodeó ninguna celda/fila/letra en producción (T0 = evidencia, nunca contrato).
5. No se tocó la plantilla (huella v4 intacta; cero `<f>`, cero insert/delete).
6. No se debilitó ningún test (0 Skip, 0 borrados, 0 tolerancias ampliadas, 0 tautologías nuevas).
7. No se hizo ningún commit/push/checkout/reset (los hace el Ingeniero).
