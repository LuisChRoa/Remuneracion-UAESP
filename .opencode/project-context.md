# Project Context — Remuneración Quincenal UAESP

> Datos específicos del proyecto. Las skills son GLOBALES (`~/.config/opencode/skills/`) y contienen SOLO metodología. Este archivo contiene los datos de proyecto. El orquestador inyecta AMBOS al delegar.

## Stack
| Capa | TargetFramework | Paquetes |
|------|-----------------|----------|
| `Remuneracion.WinForms` (UI) | `net10.0-windows` | Serilog 4.4.0 + Serilog.Sinks.File 7.0.0 |
| `Remuneracion.Infrastructure` | `net10.0` | ExcelDataReader 3.9.0 (read), DocumentFormat.OpenXml 3.5.1 (write) |
| `Remuneracion.Core` | `net10.0` | Serilog 4.4.0 (SOLO `LogContext`/logging en orquestadores `Procesador*`; modelo —Models/Constants/Interfaces/Exceptions— puro, sin sinks/config/I-O). ADR HU-14: justificado porque D5 lo ordenaba y la CLI (HU-15) también usará Serilog; si un día un consumidor de Core no puede llevar Serilog, se extrae la costura. |

**Tipo:** Desktop .NET (WinForms). **Persistencia:** archivos Excel (NO base de datos). **No hay test project aún.**

## Arquitectura
```
Remuneracion.WinForms (UI) ──► Remuneracion.Infrastructure (Excel I/O)
        │                            │
        └────────────────────────────┴──► Remuneracion.Core (dominio, sin deps)
```

Contratos en `Core/Interfaces/` → impl en `Infrastructure/`:
- `IRecaudoReader` (LeerR1/R2/R4) → `ExcelDataReaderRecaudoReader` (STUB)
- `IPlantillaWriter` → `OpenXmlPlantillaWriter` (STUB)
- `ICalculoRemuneracion`, `IValidador` → sin impl aún
- `ArchivoFuenteLocator` → ya implementado

## Comandos
```bash
# Build (solución dentro de Remuneracion.WinForms/)
dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx

# Run (WinExe — sin salida en consola)
dotnet run --project Remuneracion.WinForms/Remuneracion.WinForms.csproj

# Logs: Serilog escribe a ./remuneracion_log.txt (gitignored)
```
No hay test project / CI / lint config. Para tests: crear xUnit `net10.0` referenciando Core (e Infrastructure para I/O) y sumar al `.slnx`.

## Reglas de Negocio Críticas
- **NUNCA** sobrescribir fórmulas en el template Excel — pegar solo valores.
- **ASE4 (Bogotá Limpia)**: sin columna "Especiales" — chequear headers dinámicamente; si ausente usar 0.
- **SALDOS POR NOTA** y **RETRIBUCION NEGATIVA**: solo quincena 2 (`Periodo.NumeroQuincena == 2`); `AjustesSfT` = 0 en quincena 1.
- **Tolerancia:** diferencias ≤ ±0.5 entre calculado y fuente son aceptables (redondeo).
- **Filas dinámicas:** buscar por headers, NO por número de fila.
- **R4:** valores de reversión negativos en columna 4.
- **Redondear a entero** antes de escribir en CONSOLIDADO columna D (`DetRetri`).

## Cell Reference Map (CONSOLIDADO)
| Cells | Cálculo | Fuente |
|-------|---------|--------|
| D9:D13 | TOT_OPT | Recaudoporcomponente (R1) |
| D28:D32 | R2 Total Oportuno | SaldosaFavor (R2): Grand Total − SERV_ESP_K |
| D47:D51 | EXTEMP | R1 — Extemporáneo |
| D66:D70 | Reversión R4 | ReversiónPorComponente (R4) — negativa |
| D85:D89 | AJUSTES-SF-T | quincena 2 |
| D104:D108 | Total por ASE | suma |
| D109 | Gran Total | SUM(D104:D108) |

## Datos Reales — `Docs/Insumos/`
Archivos INSUMOS (NO commitados). **HU-20 (nueva organización, `Consolidado/` ELIMINADA):**
- `{periodo}/Conciliaciones/` — 5 archivos `Conjunta {prefijo}*.xlsx` / `Directa*.xlsx` (hoja `RESUMEN MES`, resuelta POR NOMBRE case-insensitive, no por posición; Plan 27). Fuente de las hojas `Recaudo *` (G1: `BuscarConciliacion` resuelve SOLO esta ruta, sin fallback). Julio trae hoja única; agosto multi-hoja (2ª/3ª posición) — leer "la primera" rompía el levantamiento.
- `{periodo}/R10_Remuneracion_AAAAMMQ.xlsx` — insumo de PERÍODO (hoja `DetRetri{AAAAMMQ}`, D9:D13 por ASE + D14 total). Oráculo de VALIDACIÓN del DetRetri calculado (G3) en AMBAS quincenas, NUNCA se escribe.
- Templates completados: `Remuneracion 202607-1 Total.xlsx` / `Remuneracion 202607-2 Total.xlsx`; plantilla canónica Q2 = `REMUNERACION 2026072/Plantilla_ Remuneracion 202607-2.xlsx`.
- Naming de fuentes por ASE: `Recaudoporcomponente_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{ts}.xlsx` (+ variantes `RerpoteDetalleSaldosaFavor_*` / `ReversiónPorComponente_*`).

**Columnas por quincena (G2-D2: quincena = dominio, nunca fuente):** el `RESUMEN MES` trae `VALOR 1°Q`/`N° REG. 1°Q` en D/E y `VALOR 2°Q`/`N° REG. 2°Q` en F/G. `LeerRecaudosEmpresa(periodo, …)` lee D/E si `Periodo.NumeroQuincena==1` y F/G si `==2`; las claves de `RecaudoEmpresaInputs.Celdas` reflejan la columna destino real.

**Validación DetRetri-vs-R10 (G3-D1, OBLIGATORIA en AMBAS quincenas):** `IDetRetriR10Reader` es dependencia obligatoria del constructor de `ProcesadorPeriodo` (7.º parámetro) — forma parte del flujo normal en Q1 y Q2. El DetRetri calculado bottom-up (`DetRetriQ2.Detalle = ROUND(D104:D108,0)`) se contrasta contra el R10 con tolerancia ±0.5 post-redondeo; en Q2 se escribe (`DetRetri2026072!D9:D13`), en Q1 solo se valida (la hoja `DetRetri2026071` queda con sus fórmulas protegidas intactas). Divergencia → fail-fast con período + archivo + ambos valores. R10 ausente → `No se encontró R10_Remuneracion_{AAAAMMQ} en '{carpetaPeriodo}'` (G3-D2).

**Alcance Q2 acotado (HU-20-T0b/G2-D1):** el layout R4-por-empresa Q2 sigue divergiendo (ASE2-Q2: ENEL+OCCIDENTE, sin `NUEVO ESQUEMA`), por lo que la conciliación por empresa en Q2 se OMITE (follow-up con su propio T0); las hojas `Recaudo *` SÍ se levantan en Q2 (layout `RESUMEN MES` uniforme).

## Doctrina Espejo Estructural R1 (Plan 21 — vigente)

**Principio rector (`la fuente del período define la forma`):** la hoja `Reporte Componentes R1` deja de gobernarse por mapas absolutos con ocurrencias congeladas y pasa a un **espejo estructural fila-a-fila**: la secuencia observada de `Recaudporcomponente` (firma A–E + valores por encabezado de columna) ES la especificación del período. Se eliminan las cardinalidades exigidas y las ocurrencias congeladas (`D-B`): una fila ausente en la fuente se **suprime** (nunca 0 simulado); una fila presente con valor 0 se escribe 0. Fuera de los bloques espejo, el fail-fast "slot ausente ≠ 0" sigue vigente. Evidencia base: `plans/21 - T0 Evidencia.md` (espejo 1:1 confirmado 10/10 julio; conteos Q1 38/52/46/69/44 y Q2 45/75/43/73/52).

**Derogación puntual D-C (prohibición insert/delete):** los planes 07/16 prohibían insertar/borrar filas del template ("si no alcanza, fail-fast honesto"). El Plan 21 **deroga esa prohibición SOLO para los bloques espejo R1** (y R4 según veredicto T0b), con **reanclaje obligatorio** de referencias A1 y gates de evidencia. Fuera de esos bloques, la prohibición intacta. Implementación: `OpenXmlEspejoR1Mutador` (Infrastructure) procesa los bloques **5→1** (de abajo hacia arriba), inserta/borra filas preservando estilos y fórmulas, y reanclaja fórmulas, `ref` de shared formulas, celdas combinadas y nombres definidos. Guard anti-fórmula vigente: una celda de valor/fórmula destino que sea fórmula → `ERR-PLANTILLA`, nunca sobrescritura silenciosa.

**Tabla de deltas por ASE (T0a, plantilla Q2→fuente agosto):** `ConteoAgosto − plantillaQ2`.
| ASE | Plantilla Q2 | Fuente agosto | Δ |
|-----|-------------:|--------------:|--:|
| 1 PROMOAMBIENTAL | 45 | 42 | −3 |
| 2 LIME | 75 | 66 | −9 |
| 3 CIUDAD LIMPIA | 43 | 37 | −6 |
| 4 BOGOTA LIMPIA | 73 | 79 | +6 |
| 5 AREA LIMPIA | 52 | 60 | +8 |

**Invariantes duras de cierre (T0e):** `A='Componente' B='Total'`, `A='Subs/Cont' B='Total'` y `A='Total' B` vacío — presentes en 15/15 (5 ASE × Q1/Q2/agosto). `Mes` y `AFaseo` NO son invariantes (la forma la define la fuente). Si falta una invariante dura → fail-fast que nombra ASE + reporte + fila esperada (`BloqueEspejoAseInputs.FaltantesInvariantesDeCierre`, fuente única).

**Columna L-menores absorbida:** el path legado rol/ocurrencia (`LeerLEspecialesMenores` + `WorkbookLeafCellMapInterventoria.LMenoresPorAse/Q2` + `RolLMenor` + `ObtenerLMenores`) se **retiró** (T5) al probarse absorción 15/15 (`EspejoR1AbsorcionTests`). Su columna (`SERVICIO ESPECIALES` del template) la escribe el espejo por encabezado en toda la secuencia. `WorkbookLeafInputs.LEspecialesMenores` y `LEspecialesMenoresAseInputs` eliminados. La parte INTERVENTORIA (`D2b`, hoja intacta) del mapa sigue vigente.

**Follow-up explícito (pendiente):** el **bloque espejo solo se implementó para R1** (`OpenXmlEspejoR1Mutador`). `Reversion Pagos R4` quedó declarado "ENTRA" por T0b (espejo de fila A/B/C 1:1 en 10/10) pero su implementación es **pendiente** (mismo motor, otra hoja). Otros follow-ups: R4-por-empresa Q2 (recorte HU-20/G2-D1), cobertura de `conditionalFormatting`/`dataValidations` (W-4), y el blocker de `SaldosaFavorAplicadosPorNotas` de agosto (HU-11/Q2: ASE2 sin la fila `Vlr Intereses`; fuera del scope del Plan 21).

## Doctrina Roles R1-Q2 por firma sobre la secuencia espejo (Plan 25 — vigente)

**Principio rector (`la fuente define la forma` extendido a la LECTURA de roles):** el mapa leaf R1-Q2 (`WorkbookLeafCellMapQ2.R1Q2EditablesPorAse`) deja de resolver las celdas destino por **enésima ocurrencia de roles congelados** (`ElementAtOrDefault` sobre listas filtradas por raw) y las resuelve por **firma de etiquetas (A/B/C/D/E)** dentro de la **secuencia espejo ya cargada** (`LeerEspejoR1` / `BloqueEspejoAseInputs.Filas`). El mapa declara celdas → rol; la firma y el orden de aparición resuelven el valor. **Fin del conteo de ocurrencias en ese mapa** (mismo patrón que falló dos veces: L-menores HU-12 y `Aplicacion` agosto).

**Corrección de rol D-A:** `F37` (ASE1) y `F270` (ASE3) estaban bindeadas a una fila `Subsidio` (`Subs0`) cuando la celda destino es una fila `Aplicacion`; el defecto quedó enmascarado por la coincidencia numérica de julio (`Subs0 == Aplic1` en bloques monocompañía). Ahora son `Aplic1`; el miembro `Subs0` fue **retirado** del enum (sin consumidor Q2). La resolución por firma elimina la clase de defecto (el valor sale de la fila con la firma correcta, no del rol declarado).

**Regla de ausencia por rol (T0e §6):** `Mes0/1/2` y `Lmes0/1/2` son **core OBLIGATORIOS** → fail-fast que nombra ASE + reporte + celda (mensaje intacto); `Aplic0/Aplic1/LAplic0` son **OPCIONALES** → ausente = **0 explícito** (no "no escribir": el agregado `extemp` vive en `WorkbookLeafInputsR1` y alimenta `TotalD104`/DetRetri en ambos flujos single-ASE y 5-ASE). NO se generaliza "ausencia de cualquier rol = 0": `Mes` (totOpt) y los gates de coherencia exigen el dato.

**Agregado EXTEMP (D-C, R-F-2):** `EXTEMP = ΣF(TODAS las filas Aplicacion) − Especiales de la primera` (0 si no hay ninguna). Coincide con `Aplic0+Aplic1` con 2 filas (validado 10/10 vs el extemp implícito del R10) y da 0 en ASE3-agosto (0 filas, implícito 0,33 ±0.5). Fuente única: `BloqueEspejoAseInputs.SumarAplicacion`.

**totOpt generalizado (D-F, forzado por el DoD del Plan 25):** `totOpt = ΣF(todas las filas Mes) − Especiales de todas MENOS la última`. Es idéntico al mapa congelado en ASE1-4 (3 filas) y ASE5-julio (2 filas, Lmes1 = 0), y generaliza ASE5-agosto (3 filas Mes — T0a §2.1) al mismo 5-término visible del template, sin depender del conteo. Guardián: ASE5-julio = 12.033.011.685,71 (idéntico al golden D13); ASE5-agosto = 12.105.458.586,04.

**Estados finales (Q2):** ASE1/2/4/5 conservan la rama julio (todas las filas presentes); ASE3-agosto es la rama tolerante (0 filas `Aplicacion`). La escritura espejo R1 (`omitirR1`), las protegidas `F343/F345`, los puentes AJUSTES/CONSOLIDADO y Q1 (`MapearR1`, `AjustesSfT = 0`) quedan **intactos**; el fix vive en el reader compartido (`ProcesadorPeriodo` y `ProcesadorRemuneracion`).

**Trampa documentada (R-TRAMPA-OPC):** aplicar opcionalidad (`Aplic` ausente = 0) **SIN** la corrección de rol D-A desplaza el fallo al DetRetri-vs-R10 con error ~1.088e9 en ASE3-agosto. La corrección de rol y la opcionalidad van **juntas**.

**Follow-ups vivos (fuera del Plan 25):** R4 espejo pendiente del Plan 21 (motor solo R1); R4-por-empresa Q2 (recorte HU-20/G2-D1); CF/DV del reanclaje (W-4); `R-EXTRA-CONCEPTO` del Plan 23. R-AGREGADO-N: con 1 o 3+ filas `Aplicacion` la sumatoria D-C queda validada por decisión sin evidencia runtime (su propio T0 con fuente real si cambia la semántica).

## Doctrina Preflight de insumos del período (Plan 26 — vigente)

**Principio rector (`validar TODO antes de procesar`):** ante un período con insumos incompletos, el flujo aborta en segundos —antes de abrir cualquier workbook— con UN solo error `ERR-FUENTE-NO-ENCONTRADA` (salida 2) que enumera **de una vez TODOS los faltantes o inválidos**, en vez de procesar 5 ASE para descubrir al final un faltante por corrida. El preflight es una puerta de **existencia + firma en filesystem**: NO valida contenido (headers/estructura/coherencia numérica siguen fallando en sus readers como hoy).

**Doctrina firma (Plan 27 — vigente):** además de la existencia, el preflight valida la **firma binaria** de CADA path que ya resuelve (R1/R2/R4 + banco + balance siempre; +saldos/retribución solo Q2; 5 conciliaciones; R10), con lectura acotada de ≤512 B por archivo (`InspectorFirmaExcel`, Core BCL puro — sin ExcelDataReader/OpenXML ni depender de Infrastructure). Regla mínima: PK (`50 4B 03 04`, xlsx) u OLE (`D0 CF 11 E0`, xls legacy) = válido; texto con marcadores web (`<`, `<!doctype`, `<html`, `MIME-Version`, `Saved by`, `From:`) = *"parece una copia de una página web"*; resto = *"no es un Excel válido"*; error de I/O = se trata como desconocido (nunca throw propio). Un archivo inválido ocupa el lugar del archivo en la MISMA lista numerada y MISMO código `ERR-FUENTE-NO-ENCONTRADA` (el usuario ve un formato, no dos). El detalle técnico (firma + primeros bytes hex) va SOLO al `Log.Warning`; el mensaje al usuario es administrativo (prohibido "magic bytes", "MIME", "firma binaria", "header", "MHTML", "0x"). Finders, readers, cálculo y escritura intactos: con insumos válidos el flujo es bit-idéntico (cero-geometría). El preflight NO valida nombres de HOJA (abriría workbooks: sigue fuera).

**Doctrina hoja-por-nombre (Plan 27 — vigente):** las conciliaciones se leen con `ExcelWorksheetNavigator.LeerFilas(ruta, "RESUMEN MES")` (match case-insensitive, fail-fast que nombra archivo+hoja), nunca "la primera hoja a ciegas". Julio (hoja única) es bit-idéntico; agosto (multi-hoja) lee la hoja correcta. La quincena sigue gobernándola el dominio (`Periodo.NumeroQuincena`).

**Superficie verificada = superficie que el runtime resuelve (ni un archivo más ni uno menos):** por ASE, R1 + R2 + R4 + banco + balance SIEMPRE; `SaldosaFavorAplicadosPorNotas` y `RetribuciónNegativa` SOLO si `Periodo.NumeroQuincena == 2` (gobierno por dominio, nunca por detección de contenido). De período: las **5 conciliaciones SIEMPRE (Q1 y Q2)** — la lectura `Recaudo *` está fuera de toda rama por quincena (§V3) — + R10 SIEMPRE (oráculo obligatorio G3 en ambas quincenas). `RecaudosReversados` queda **FUERA**: el runtime no lo consume (sin finder ni lectura, §V9); exigirlo sería inventar un requisito.

**Finders intactos; el preflight valida, no resuelve (D-C):** `ValidadorInsumosPeriodo` REUTILIZA los finders de `ILocalizadorArchivosAse` (mismo prefijo case-insensitive, misma normalización de diacríticos, mismo `TopDirectoryOnly`); no duplica lógica de match ni reemplaza la resolución runtime. Los `?? throw` intermedios de los procesadores **permanecen** como defensa en profundidad. Si UAESP agrega o retira un insumo que el runtime resuelve, se agrega/retira en la lista del validador con su finder (checklist de paridad preflight↔runtime).

**Mensaje en lenguaje administrativo (D-E):** encabezado `Faltan insumos para el período {AAAAMMQ} (quincena {1|2}). No se procesó ningún ASE.` + lista numerada; cada ítem dice QUÉ falta (nombre que el usuario reconoce: ASE, reporte, carpeta), DÓNDE debe ir (ruta y nombre exacto o inicio de nombre esperado) y QUÉ hacer (solicitar/generar, colocar, reejecutar). Los archivos **inválidos** se suman como factorías propias (Plan 27): *"El archivo 'X' no es un Excel válido (parece una copia de una página web)"* o *"...no es un Excel válido."* + ubicación + acción; faltantes e inválidos conviven en la MISMA lista con numeración única y orden estable (el ítem inválido ocupa el lugar del archivo). **Jerga de código PROHIBIDA** en el texto de los ítems (`prefijo`, `matcher`, `finder`, `TopDirectoryOnly`, `magic bytes`, `MIME`, `firma binaria`, `header`, `MHTML`, `0x`, nombres de clase o código interno). Si falta la carpeta `Conciliaciones/` completa, UN solo ítem la nombra + enumera los 5 archivos esperados (no 6 líneas redundantes); igual para una carpeta ASE ausente. Orden estable: ASE 1..5 (carpeta, luego reportes en el orden R1, R2, R4, banco, balance, [Q2] saldos-notas, retribución), luego `Conciliaciones/`, luego R10.

**Integración:** `ProcesadorPeriodo.Ejecutar` corre el preflight tras resolver las 5 carpetas ASE y ANTES del loop por ASE (`progreso?.Report` + `Log.Warning`; doctrina HU-14); con faltantes NO se lee, no se calcula y no se escribe. `ProcesadorRemuneracion.Ejecutar` (single-ASE) solo verifica existencia (`File.Exists`) de R1/R2/R4 — guardrail mínimo; NUNCA exige banco/balance/conciliaciones/R10/saldos. Sin faltantes el flujo es bit-idéntico al actual (cero-geometría: ningún valor ni fórmula cambia; el diff es el servicio/modelo Core + 2 puntos de llamada).

**Follow-up explícito (fuera del Plan 26):** validación de CONTENIDO anticipada (headers/estructura antes de procesar) requiere su propio T0: abre workbooks, bloquea archivos y duplica readers (descartado por diseño, no por olvido). El Plan 27 cubrió la **firma binaria** (bytes, sin abrir workbooks) y la **hoja por nombre** en runtime; la validación de NOMBRES de hoja en el preflight sigue como follow-up (abriría workbooks).

## Doctrina Integridad del workbook de salida + sello de fechas (Plan 28 — vigente)

**Unidad S — saneamiento de la cadena de cálculo (`calcChain`).** Al guardar, el writer **elimina `CalculationChainPart`** (tolerante a plantillas que no la traen) y setea **`FullCalculationOnLoad=true`** (`<calcPr fullCalcOnLoad="1"/>`; `calcId` preservado), vía el helper único **`SaneadorCadenaCalculo.Sanear(WorkbookPart)`** invocado en los 3 caminos de guardado de producción: `OpenXmlPlantillaWriter.GenerarWorkbook` single-ASE (L147) y multi-ASE (L287), y `OpenXmlEspejoR1Mutador.Ajustar` (espejo standalone, L85). **NUNCA toca `<f>` ni `<v>`**. Motivo (evidencia T0): el espejo R1 mueve filas pero la calcChain de la plantilla se preserva → salida de agosto con **1003 entradas calcChain inconsistentes** (12921 entradas / 12921 fórmulas) → Excel muestra el diálogo de reparación al abrir. Con el saneamiento la salida abre sin reparación y Excel recalcula una vez al abrir. Goldens ±0.5 intactos por construcción (julio Δ=0 → sello idempotente). El 3.er `Save()` (`Herramientas/VerificadorRecaudo`) queda FUERA (herramienta de diagnóstico, no camino del writer).

**Unidad F — sello de fechas del período.** `ProcesadorPeriodo` lee el R10 del período (`DetRetri{AAAAMMQ}`: `G7`=Fecha Desde, `J7`=Fecha Hasta), valida que sean legibles (**fail-fast `ERR-FORMATO-FUENTE`** nombrando período + archivo R10 + celda si llegan en `default`) y las propaga a `ResultadoRemuneracion.FechaDesde/FechaHasta`. El writer sella **`CONSOLIDADO_TOTAL RECAUDO!G7 ← FechaDesde`** y **`!K7 ← FechaHasta`** como seriales OADate, **solo valores** sobre celdas existentes (estilo fecha `s=9`, numFmt 14, preservado), con el guard anti-fórmula existente como red (celda-fórmula → `ERR-PLANTILLA` nombrando hoja+celda). **Dato T0:** la columna REAL de "Fecha Hasta" en el template es **K7**; `J7` es el rótulo (con el typo de la fuente "Feha Hasta:"). El path single-ASE sin R10 no sella (no-op). C59 (quincena, hoja banco) sigue gobernado por dominio y NO se duplica; `N3` (código remuneración) y `C6/D6` (sello de proceso) quedan fuera de scope (no son el rango del período).

**Regla de proceso — SIEMPRE plantilla en ceros.** Se corre desde la **plantilla en ceros** (`REMUNERACION 2026072/Plantilla_ Remuneracion 202607-2.xlsx`), **NUNCA** desde una salida ya trabajada de otro período: la salida trabajada arrastra fechas y cadena de cálculo ajenas (el defecto que Plan 28 cierra). Veredicto T0/D-G: **sin guardrail de bloqueo** (documentación + regla de proceso; un detector "salida trabajada" arriesgaría falsos positivos). El sello hace que cada período escriba SU rango desde su R10.

## Doctrina de paridad app-vs-manual (Plan 29 — vigente)

**Principio rector (`lo que el workbook puede derivar de los insumos, la app lo escribe; lo que es proceso, se documenta`):** la salida de la app debe ser indistinguible del archivo manual del administrativo en todo lo que el workbook puede derivar de los insumos. El Plan 29 cierra la brecha con un T0 bloqueante que arbitró cada categoría en disco: **no hay una única causa** (D-A). Cero fórmulas tocadas (`<f>` intacto en todo el diff) y DetRetri-D 5/5 vs R10 intacto por construcción.

### Tabla de causas V2..V7 (veredicto T0)

| # | Hallazgo en disco | Tipo | Consecuencia |
|---|---|---|---|
| V2 | Agregados R2/R4 idénticos app vs manual al centavo (`E17/E28/K17`, `D15/P15`; post-recálculo `E43=E17+E28-K17`, `D73=D15-P15`) | **No-brecha** — no se toca | El encargo comparó celdas Q1 literal-0 y cachés `<v>` stale pre-recálculo; el writer SÍ usa el mapa Q2 |
| V3 | Recaudo por empresa: columnas de 2.ª quincena (F/G) idénticas; las de 1.ª quincena (D/E) quedan en 0 en una corrida Q2 aislada | **No-brecha por diseño** — no se toca | Cada quincena escribe SOLO sus columnas (`Periodo.NumeroQuincena`: D/E en Q1, F/G en Q2). Comparar D/E es comparar quincenas distintas (el manual acumula Q1+Q2) |
| V4 | BCE ASE1-4: swap D/E puro (F=D+E idéntico). **Corregido en T1** | Causa 1 | Header manda: template-D = SUBSIDIO ← columna E-fuente; template-E = CONTRIBUCION ← columna F-fuente. Refuta el T0-0.2 del Plan 10 («D=CONTRIBUCION»); F conmutativa → el golden ±0.5 no detectaba el swap |
| V5 | BCE julio-ASE5: los `D7/E7` del manual no aparecen en ningún insumo (48 xlsx + 40 pdf, 0 apariciones) | Causa 2 — **divergencia-del-manual** | La app es fuente-fiel (`D7/E7` == `TOTAL GENERAL` E39/F39 al centavo). NO se toca `LeerBalanceSc` por ASE5; se documenta (T0a / R-B-2 / S2) |
| V6 | Detalle R2/R4 (filas `E3:R3` y `D3/D9`-style) en 0 en la app aunque la fuente lo trae fila a fila | Causa 3 — **Unidad R** | Se lee por LABEL (filas) + ENCABEZADO (columnas) y se escribe en la misma pasada atómica; `Especiales` ausente = 0 explícito. Sin insert/delete (las filas ya existen en ceros) |
| V7 | `DetRetri`/`DetValiRetri` C..O en 0 en la app; el manual los trae literales | Causa 4 — **Unidad D** | **12/12 DetRetri y 11/12 DetValiRetri** con origen workbook-interno (`CONSOLIDADO_TOTAL RECAUDO` fila `104+k`) se escriben como literal redondeado. **SALE: `DetValiRetri!J9` (`AJUSTE A LA DECENA`, manual-externo)** con motivo |

**Contexto V1/V8:** la plantilla trae las hojas leaf como literales-0 de captura y los consolidados como fórmulas (la app solo escribe literales; todo lo demás recalcula solo). `V8` (`VALIDACION_TOTAL!C15` False) se explica por el literal faltante `'Valida - Control Recaudo'!F10` (T0f), que la Unidad P puebla.

**Comparador de regresión — exclusiones versionadas (D-E):** `ComparadorSalidaVsManual` (Infrastructure, BCL puro `System.IO.Compression` + `System.Xml`; sin Excel/COM/ExcelDataReader) demuestra paridad por (a) literales ±0.5 o igualdad de texto, y (b) **igualdad de texto de fórmula** (misma fórmula + mismos inputs ⇒ mismo resultado post-recálculo; `fullCalcOnLoad` garantizado por Plan 28). **Prohibido depender de cachés `<v>`** (stale por diseño: V2/V8). Exclusiones declaradas y versionadas (R-FALSO-POSITIVO): columnas de la otra quincena (V3), metadatos heredados de período (`N3`/nombres de hoja salvo proceso — H1) y `DetValiRetri!J9` (SALE de Unidad D).

**Proceso por quincena (D-F, H1):** cada quincena corre desde **su base propia del período** (plantilla en ceros + R10 del período). Para acumulado: Q1→Q2 encadenado en el mismo workbook. Para comparar por quincena: bases separadas (una corrida por quincena desde la plantilla en ceros). Correr Q2 sobre la salida de otro período arrastra el sello de período (`N3`) y los nombres de hoja (H1: la salida de agosto heredó `2026072`), aunque `G7/K7` sí se sellan por corrida (Plan 28). **Sin guardrail de bloqueo** (doctrina Plan 28 D-G).

## Doctrina de nombres de hoja dinámicos por período + base canónica (Plan 30 — vigente)

**Principio rector (`el nombre de hoja se COMPONE del dominio, no se detecta`):** los nombres de hoja que dependen del período (`DetRetri{AAAAMMQ}`, `DetValiRetri{AAAAMMQ}`, `Informe AFaseo Recaudo {AAAAMM}-{Q}`) se resuelven desde el dominio con el helper puro de Core **`NombresHojaPeriodo`** (`Remuneracion.Core/Models/NombresHojaPeriodo.cs`): `DetRetri(codigoCompleto)` → `DetRetri2026082`; `DetValiRetri(codigoCompleto)`; `InformeAFaseo(codigoAAAAMM, quincena)` (con sobrecargas `(Periodo)`). **El sufijo ES `Periodo.CodigoCompleto`** (2026071, 2026072, 2026082). Prohibido resolver por enumeración de hojas del workbook ("la que empiece por `DetRetri`"): la detección por contenido es frágil ante hojas heredadas de otro período (defecto H1). Cero literales `2026xxx` de nombre de hoja en runtime de producción fuera del helper.

**OCP hacia el futuro:** un período nuevo funciona con **cero cambios de código** — el mismo helper resuelve `DetRetri2026091` (septiembre 2026091) sin tocar nada; Q1/Q2-julio resuelven byte-idéntico a los literales históricos (regresión por goldens ±0.5). El sustituto Q1→Q2 de fragmentos "Protegidas" (`ProtegidasBceParaPeriodo`) aplica `Replace(ancla-Q1, periodo.CodigoCompleto)`, nunca un literal de quincena.

### Base canónica por período (D-B/D-E)

**Regla de proceso:** cada período corre desde **su base canónica** (plantilla en ceros) que trae las hojas `DetRetri{código}` / `DetValiRetri{código}` / `Informe AFaseo Recaudo {AAAAMM}-{Q}` **del período** e internamente consistente con el manual del administrativo (metadatos `N3`/fechas + columnas Q1 D:E heredadas del manual tal cual, sin recalcular).

| Período | Base canónica |
|---|---|
| Julio (2026072) | base de julio (consistente, intacta) |
| Agosto (2026082) | `Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx` — nombre propio (R-BASE-DOBLE); ejemplo de la doctrina |

**Herramienta de preparación (offline, one-shot):** `Herramientas/PreparadorBasePeriodo/` (BCL puro `ZipArchive` + `Regex`, **fuera del runtime/pipeline**) convierte una copia de la plantilla base en la base canónica del período: renombra los 3 `<sheet name>`, reescribe el token de período en las 11 `<f>` que lo referencian, actualiza `docProps/app.xml` y copia metadatos/Q1 del manual celda por celda (allow-list, solo destinos sin fórmula). **El runtime NUNCA toca `<f>`** (invariante Plan 28/30): la reescritura mecánica de fórmulas corre UNA vez, offline; cero fórmulas de negocio cambiadas.

### Fail-fast de hoja ausente (D-C)

**Principio:** si la plantilla no trae la hoja esperada del período → `ERR-PLANTILLA` nombrando **hoja esperada + período + operación** (p. ej. `La hoja 'DetRetri2026082' no existe en el workbook para escritura DetRetri-Q2 (período 2026082).`). **Sin fallback** a hojas de otro período: un fallback silencioso escribiría el dinero del período en la hoja del período equivocado. Aplica en los 3 puntos de resolución de hoja Det (escritura Q2, Unidad D trazable, oráculo `DetValiRetri`); base equivocada → error accionable inmediato, sin escritura parcial.

**Informe AFaseo — cero código (D-D):** el runtime ignora `Informe AFaseo Recaudo` por spec (no se lee ni se escribe; ninguna `<f>` lo referencia). Solo existe como hoja en la base; su nombre lo fija la doctrina de naming, pero no hay consumidor en código.

## Doctrina de recomposición por firma de totales R1 + gate workbook-vs-dominio + sello de proceso (Plan 31 — vigente)

**Principio rector (`las fórmulas visibles del espejo se COMPONEN por firma, no se reanclan por conteo; los gates leen el WORKBOOK, no solo el dominio`):** los totales visibles de `Reporte Componentes R1` (TOT_OPT `F`, total TDF `G` y EXTEMP `F` por ASE) se recomponen en un **pase final post-5→1** del mutador (**D-A**: el mutador posee la geometría final; el writer con `omitirR1` no la conoce) desde los roles reales de la fuente (firma, no cardinalidad ni borde). Es la única excepción legítima y acotada al invariante "NUNCA sobrescribir fórmulas" (D-D).

### Causa cerrada (T0 — no re-investigar)

El espejo estructural R1 **dimensiona por conteo** (`delta`) y **borra en el PIE** del bloque (`primeraBorrada = totalRowIdx - m`), reanclando solo `r >= primeraBorrada`; pero la fuente de agosto **recorta en la CABEZA** (Mes 12/32/48 → 9/29/45). Los DATOS se escriben bien por orden, pero las `<f>` de los visibles quedan ancladas a las filas de julio (`F50 = F32+F47+F12-L12-L32` en vez de `F29+F45+F9-L9-L29`) y nadie re-deriva las filas-ancla por firma. **ASE5 exige recomponer** (no reanclar): su fuente trae **3 filas `Mes`** (agosto) mientras la plantilla declara 2 términos — un mapa de filas no puede agregar un término (D-B). Era un defecto de **anclaje de `<f>`**, no de lectura ni de dominio: los datos app-vs-manual son idénticos fila a fila y el dominio C# ya calculaba `totOpt`/`extemp` correcto por firma.

### Excepción `<f>` (D-D) — alcance exacto

- Se reescriben **solo** las celdas visibles del contrato T2: `TOT_OPT F` = `ΣF(Mes) − ΣL(Mes menos la última)`; total TDF `G` = `ΣG(Mes)`; `EXTEMP F` = `ΣF(Aplic) − Esp(primera)` (0 Aplic → literal `0` auditado). Forma idéntica al dominio `MapearR1Q2`.
- **Fuera de esas celdas el invariante sigue intacto:** el `Reanclar` mecánico se conserva para refs externas (`CONSOLIDADO!D9='R1'!F50`), nombres y merges; R2/R4-detalle, CONSOLIDADO y validaciones no se tocan.
- **Julio-identidad (D-E):** con geometría de julio (delta 0) la recomposición reproduce byte-idéntica la fórmula canónica (`F53=F32+F48+F12-L12-L32`, …); si julio difiere, es regresión.
- Auditoría término a término: texto de fórmula vs manual (f-vs-f) + evaluación BCL vs dominio ±0.5 (el gate).

### Gate workbook-vs-dominio R1 (C, S1)

`ValidadorTotalesR1Workbook` (Infrastructure, BCL/OpenXML en lectura — sin Excel/COM) lee el workbook **generado**, evalúa cada visible (texto-`<f>` + literales; **prohibido `<v>` stale**) y lo contrasta contra el dominio por firma (`TotalOportunoEsperadoPorAse`/`ExtemporaneoEsperadoPorAse`) ±0.5, con fail-fast que nombra ASE+celda si una ref no resuelve. Cubre el punto donde la validación de protegidas se **salteaba** con `espejoDesplazado=true`. **S1:** julio PASS siempre (identidad); agosto FAIL en el estado pre-T2 (reproduce el bug) y PASS tras T2. En paralelo, el comparador de regresión pierde la exención *blanket* de R1-agosto: `Reporte Componentes R1` vuelve a compararse f-vs-f (solo exclusiones versionadas con cita).

### Sello de proceso D6/D7 (D-F) — texto, no serial

`ProcesadorPeriodo` lee del R10 (`DetRetri{AAAAMMQ}`) **`Fecha de Proceso` → `FechaProceso`** y **`Hora` → `HoraProceso`**, con fail-fast `ERR-FORMATO-FUENTE` (período+archivo+celda) si ilegibles; el writer sella **`DetRetri{AAAAMMQ}!D6/D7`** y **`DetValiRetri{AAAAMMQ}!D6/D7`** (nombres por `NombresHojaPeriodo`) con guard anti-fórmula. **Desviación ratificada por el mini-T0 (R-S-1):** el plan D-F decía "valores OADate"; las 4 celdas traen formato **Texto** (`numFmtId=49`) y la base/manual/R10 guardan el sello como **texto** (`dd/MM/yyyy` + `hh:mm AM/PM`) — un serial OADate mostraría el número crudo. Se sella **texto** preservando el estilo (julio `04/08/2026` `10:15 AM`; agosto `02/09/2026` `07:42 AM`). Distinto del sello de **rango** del Plan 28 (`CONSOLIDADO G7/K7`, serial OADate). **`CONSOLIDADO!D6` NO se duplica** (ya sellado en la base; R-S-3).

### Veredictos R2/R4+soporte (T4 — solo documenta, cero código de producción)

| ID | Frente | Veredicto | Consecuencia (comparador / código) |
|---|---|---|---|
| R-D-1 | R2-detalle (`Rem. Anticipos R2`) | declarar-divergencia | Ya versionado (brecha ítem 9); código: nada |
| R-D-2 | R4-detalle (`Reversion Pagos R4`) | declarar-divergencia | Ya versionado; código: nada (idem R-D-1) |
| R-D-3 | SALDOS POR NOTA (detalle) | declarar-divergencia | Versionar la categoría (agosto); código: nada; re-encender al cerrar la malla (F-T4-2/3) |
| R-D-4 | AJUSTES - SF-T + arrastre (`CONSOLIDADO_TOTAL RECAUDO`, `REMUNERACION_*`) | declarar-divergencia | Arrastre de R-D-1/2/3; versionar la categoría |
| R-D-5 | R1-interior subvisible (`Reporte Componentes R1`, F/G/H interiores) | adoptar-geometría | NO versionar: los 96 quedan en ROJO como brecha viva; follow-up F-T4-1 |
| R-D-6 | INTERVENTORIA R26 | declarar-divergencia (cosmética) | `=F15` vs `=H15`, mismo valor 189185988; versionar acotado a R26 |
| R-D-7 | ANT EXT-REV / ANTICIPOS USUARIOS | fuera-de-alcance | Ya versionado (non-goal del Plan 29); sin cambio |

**Follow-ups (fuera del Plan 31):** F-T4-1 (R-D-5: extender la recomposición por firma al INTERIOR de R1 — Plan 32, T0 propio); F-T4-2 (R-D-1/2: remesh de filas R2/R4 si el Ingeniero autoriza insert/delete o mapeo por label); F-T4-3 (R-D-3/4: re-encender o promover la exclusión a permanente); F-T4-4 (R-D-6: confirmar si `R26` es cosmética del manual). Tras las exclusiones temporales de T4, agosto pasa de 1272 → 96 divergencias inesperadas (todas R-D-5).

**Lección (método):** los gates de coherencia eran **dominio-a-dominio** y el comparador eximía R1 en bloque → un total mal anclado se entregaba en silencio. Regla que queda: **un gate debe leer el WORKBOOK generado** (texto-`<f>` + literales), no solo los agregados C# del dominio.

## Estructura de Directorios
```
Automatización/
├── README.md                       # Spec de negocio completa — leer primero
├── Docs/ (docs + Insumos/)
├── Remuneracion.Core/
├── Remuneracion.Infrastructure/
├── Remuneracion.WinForms/          # contiene el .slnx
├── plans/                          # planes SDD (fuente de verdad)
├── requirements/                   # requerimientos formales
└── .opencode/ (agents/, project-context.md, opencode.jsonc)
```

## Fase Roadmap
1. **Fase 1 ✅ cerrada (HU-01..HU-06):** prototipo single-ASE — lectura R1/R2/R4, motor CONSOLIDADO, escritura OpenXML, UI, golden Capa A.
2. **Fase 2 ✅ cerrada (HU-07..HU-13):** 5 ASE, conciliación (2.2), banco (2.3), BCE (2.4), AJUSTES Q2 (2.5), cell-map Q2 + DetRetri (2.6), validaciones cruzadas (2.7).
3. **Fase 3 (HU-14..HU-17, aprobado por el Ingeniero 2026-09-09):**
   - HU-14: 3.1 + 3.2 robustez y observabilidad (cierre errores + logging; absorbe W-1/W-2/W-3, S-1..S-4).
   - HU-15: 3.3 modo CLI (ejecución desatendida; habilita pruebas repetibles).
   - HU-16: INTERVENTORIA L25:N31 + filas L-Especiales menores (última hoja sin HU, con su T0).
   - HU-17: 3.4 manual + entrega a pruebas (manual, instructivo Capa B, paquete, casos, criterio de pase).
   - Acción del Ingeniero (no HU): Capa B manual Excel con el instructivo de HU-17.

## MCPs y Fuentes de Contexto
- `engram` — memoria persistente (decisiones, continuidad, aprobación de planes).
- `context7` — documentación actualizada de librerías (ExcelDataReader, OpenXML, WinForms).
- `excel` — lectura/escritura de archivos `.xlsx` para validar datos del proyecto.
- `jira`, `playwright` — según tarea. Ninguno ejecuta cambios directos sobre BD (no hay BD en este proyecto).

## Convenciones
- Comunicación en español; dirigirse al usuario como **Inge** o **Ingeniero**.
- Finales de línea: no hay `.gitattributes`/`.editorconfig` → usar estándar del SO (Windows `\r\n`) y no mezclar.
- Commits solo con `#commit`/`#push` explícito.
- Build limpio (0 warnings) antes de marcar tareas como completas.

## Preservación de shared-formula R1 (Plan 33 — doctrina viva)

- **Doctrina (D-A/D-B):** al recomponer fórmulas de `Reporte Componentes R1`, si la celda YA trae `CellFormula` se muta **`.Text` in place** preservando `FormulaType`/`Reference`/`SharedIndex` del master shared de la plantilla; solo se crea `new CellFormula` cuando la celda nace sin fórmula. Regla por **EXISTENCIA**, no por lista de celdas: un master futuro se preserva sin código nuevo.
  - Costura única: `OpenXmlEspejoR1Mutador.EscribirFormulaPreservando(celda, texto)`, usada por `EscribirFormulaVisible` y `EscribirFormulaInterior`. `EscribirCeroVisible`/`EscribirLiteralInterior` (literal 0) **intactos**: un literal no es master.
  - Síntoma del bug que esto evita: `new CellFormula(texto)` descarta `t`/`ref`/`si` de un master → seguidoras `t="shared"` quedan huérfanas → Excel abre con «Registros quitados: Formula compartida» (Julio-2, masters `G53` si=0 y `G468` si=20).
- **Invariante workbook-wide (D-C):** en `sheet10.xml` (= `Reporte Componentes R1`), todo `si` de una seguidora (`t="shared"` con `si` y sin `ref`) tiene un master (`t="shared"` con ese `si` + `ref`) en la misma hoja. Gate de tests: `SharedFormulaR1Gate.ExigirInvariante` (+ fixture `SharedFormulaR1Esperados`: 35 masters, `si` 0..34, 1035 seguidoras). **PROHIBIDO** usar cachés `<v>` como oráculo (doctrina Plan 28/29).
- **Lección (D-G, Plan 31 → 33):** el gate debe leer el **WORKBOOK generado** (texto-`<f>` + atributos `t`/`ref`/`si` + literales), no solo el dominio C# — el dominio nunca vio el `si` huérfano.
- **Nota E7:** `G468` está también en `CeldasInterioresAgosto[4]` (`EXT_INT`), por eso el fix cubre el pase **visible e interior**.
- **Operativa:** regenerar SIEMPRE a una ruta **temp FRESCA** (nunca sobre una salida existente): `dotnet run --project Remuneracion.Cli -- --periodo <AAAAMMQ> --carpeta <insumos> --plantilla <xlsx> --salida <temp>`. Un FAIL del gate nombra hoja+celda+`si` de cada seguidora huérfana.
- **Fase 2 (T0 agosto-2, veredicto H1):** `ASE2 F199 [SUB_EMP] [Mes0]` aborta porque el rótulo del template `RECIPROCIDAD` (sinónimo legado de **EAAB Reciprocidad**) no coincide con el nombre vigente de la fuente `NUEVO ESQUEMA` (`Detalle de plantilla`: «NUEVO ESQUEMA = recaudo EAAB Reciprocidad»; `Proceso de Recaudo`: ASE2 = ENEL+EAAB+OCCIDENTE). Evidencia: `plans/33-T0-Evidencia-Agosto2.md`. El fix de agosto (T5) quedó **implementado por el Plan 34**.

## Sinonimia de empresas R1 (Plan 34 — doctrina viva)

- **Doctrina (D-A/D-B):** el rótulo de empresa del interior R1 se reconcilia por DATOS, no por ramas. Tabla `Remuneracion.Core.Models.SinonimosEmpresaR1` (clases de equivalencia citadas + matcher puro `SonMismaEmpresa`); la consume SOLO el overload `R1FirmaInterior.EsDatoEmpresa(fila, empresa)` (único punto template↔fuente del interior: cubre `SUB_EMP`/`EXT_INT`/`SUBS` sin tocar el mutador). Añadir una divergencia futura = añadir una FILA citada (sin cita no hay fila — R-SINONIMO-ABUSO); cero lógica nueva.
- **Catálogo ABIERTO (reafirmado, Plan 32 D-B):** PROHIBIDO `if (empresa == "ENEL"/"OCCIDENTE"/"RECIPROCIDAD"/"NUEVO ESQUEMA")` en producción. Un rótulo no tabulado solo se iguala a sí mismo; las empresas sin divergencia quedan identidad por construcción.
- **Lección legado↔vigente (Plan 33-T0 H1, agosto-2026082):** el template trae `C199=RECIPROCIDAD` (marca histórica de EAAB) y la fuente R1 trae `NUEVO ESQUEMA` (nombre vigente); son la MISMA EFC (EAAB Reciprocidad). Citas: `Detalle de plantilla` («NUEVO ESQUEMA: Corresponde al recaudo "EAAB Reciprocidad"…»), `Proceso de Recaudo` (ASE2 = ENEL+EAAB+OCCIDENTE), oráculo `Remuneracion 202608-2 Total_7721.xlsx` `R1!F199 = F140+F114-L114` (filas `NUEVO ESQUEMA`). NO se renombra el template: el rótulo legado es del negocio y el manual-oráculo también lo trae.
- **D-H cero-quema-de-períodos:** el matcher no recibe ningún período; la tabla es de EMPRESAS, no de períodos. Los períodos existen solo como datos de prueba. `EsPeriodoJulio` y `CeldasInterioresAgosto` intactos.
- **Normalización:** mayúsculas invariantes + sin diacríticos (misma regla que `Normalizar` del mutador); la comparación exacta case-insensitive previa se conserva como caso base (julio NO-OP por `EsPeriodoJulio`, luego el overload no se ejecuta en julio → identidad por construcción).
- **Gate (Plan 34):** `F199SinonimoTests` — repro in-process rojo→verde con insumos REALES por ruta explícita (`Docs/Prueba Agosto-2/...`; PROHIBIDO `Insumos.Raiz()`/`Docs/Insumos`, borrado por el reorg `a867706`). Fixture que manda: template 2026082 + fuente ASE2 `{ENEL, NUEVO ESQUEMA, OCCIDENTE}` + `F199=F140+F114-L114`. PROHIBIDO usar cachés `<v>` como oráculo.
- **Fail-fast enriquecido (D-G):** si la equivalencia deja 0 filas, el throw nombra AMBOS lados (rótulo-template + etiquetas fuente observadas del bloque); nunca 0 silencioso ni composición parcial.
- **Follow-ups declarados (fuera del Plan 34):** E6 — mapa ASE2 incompleto (`F209` ENERBIT / `F219` CIUDAD LIMPIA-ACUEDUCTO, clase L-1 del TODO(T2-full) Plan 32); aborto posterior desenmascarado `ASE4 F463 [EXT_INT] [Aplic0]` (R-MAPA-OCULTO; pre-existente, no regresión de T5); sucesor por-período de `CeldasInterioresAgosto`; reconciliación de rutas raíz de los 511 rojos de fixture-root (dependencia declarada). Ver `plans/34 - fix-f199-sinonimo-reciprocidad-nuevo-esquema.md`.

## Plantilla oficial directa + Phase 0 automática de base (Plan 39 — doctrina viva)

**Principio rector (`la plantilla oficial es el INPUT DIRECTO; la base del período se prepara sola`):** la plantilla oficial `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` es el input ÚNICO que consume la App para TODOS los períodos. Al arrancar, el pipeline ejecuta una **Phase 0 automática** sobre una COPIA de trabajo (el oficial NUNCA se modifica in place): renombra las 3 hojas de período (`DetRetri{AAAAMMQ}`, `DetValiRetri{AAAAMMQ}`, `Informe AFaseo Recaudo {AAAAMM}-{Q}` vía `NombresHojaPeriodo`), reescribe los 11 tokens `<f>` de período, sella `N3 = CodigoCompleto` y selecciona la columna `L9:L13` por quincena. El paso manual intermedio de **bases preparadas offline** (carpeta de bases por período) SALE del flujo de producción y se retira (ningún código lo referencia; la carpeta ya no existe en el repo).

**Invariante re-encuadrado (precisión quirúrgica):** la preparación de base (fase mecánica, explícita y auditada) SÍ puede tocar `<f>` — allow-list EXACTA de **16 celdas** (11 tokens de período + 5 `L9:L13`); el PROCESAMIENTO de negocio (slot writer values-only y todo lo downstream) NUNCA toca `<f>` ni inserta/borra filas. El grep de cierre distingue ambas fases por namespace/clase (`PreparadorBasePipeline` vs el procesamiento values-only).

**R-INTERV-Q (regla de negocio de primera clase):** `CONSOLIDADO_TOTAL RECAUDO!L9:L13` es quincena-dependiente por DOMINIO (nunca por contenido): `Periodo.NumeroQuincena == 1` → `INTERVENTORIA!N26:N30` (Primera quincena); `== 2` → `INTERVENTORIA!M26:M30` (Segunda quincena). Tercera forma = `ERR-PLANTILLA` nombrando hoja+celda+variante esperada. Oráculo de cachés por variante: `InterventoriaDeclarada.PrimeraQuincenaPorAse` / `SegundaQuincenaPorAse` (se REUTILIZA, no se duplica). Prohibido `<v>` como oráculo de fórmulas.

**Phase 0 (`Remuneracion.Infrastructure.Excel.PreparadorBasePipeline`, BCL puro `ZipArchive`+`Regex`, sin OpenXML/ExcelDataReader):** código MOVIDO (no reescrito) del preparador offline probado + el paso de selección de columna L (tabla indexada por `NumeroQuincena`; cero literales de período/empresa — D-H). Se invoca en las 3 rutas de guardado del writer (`GenerarWorkbook` single-ASE y multi-ASE, y `EscribirEspejoR1`), entre el `File.Copy` y `ValidadorBasePeriodo.Exigir`. `Herramientas/PreparadorBasePeriodo` queda como wrapper delgado (misma implementación viva; paridad byte-a-byte salvo la columna L parametrizada).

**Preflight de base (`ValidadorBasePeriodo`, extendido):** corre sobre la copia auto-preparada y exige la variante L de la quincena del run (Q1→N, Q2→M); rechaza una tercera forma con `ERR-PLANTILLA`. Resto intacto: hojas del período + `N3 == CodigoCompleto` + sellos `G7/K7/D6` + `D6/D7` + 0 tokens de período stale en `<f>`.

**Healthy-open (Plan 28, intacto):** cada salida se guarda SIN `xl/calcChain.xml` y con `fullCalcOnLoad="1"` (`SaneadorCadenaCalculo.Sanear` en los 3 puntos de guardado) — la clase de defecto "diálogo de reparación" está cerrada; el invariante shared de R1 (Plan 33) es la segunda red ciega.

**Sweep de refs duras de quincena (T3 — `Docs/T0-Plan39-Fase0-L9L13.md` §6):** barrido BCL de toda referencia con olor a quincena; cada hallazgo con veredicto citado (intencional-por-quincena con su regla, o invariante). Único hallazgo elevado sin decisión de negocio: **F-NEW-1** — familia de auto-checks booleanas que hardcodean la columna M (`DetRetri{per}!L32:L36`, `DetRetri{per}!O32:O36`, `DetValiRetri{per}!I24:I29`, `INTERVENTORIA!S26:S31`). **Medido** (Plan 39/T5, recálculo Excel sobre una corrida Q1 desde el oficial): evalúan **FALSE en Q1** para ASE 1/2/3/5 (y las filas de total), TRUE para ASE 4 (M=N para ASE4). Son celdas de auto-chequeo (no montos de negocio); requieren su propio T0 + decisión del Ingeniero — NO se implementan en Plan 39.

**Re-congelamiento:** un cambio del oficial por el negocio exige un **nuevo T0** (re-lectura BCL de AMBAS variantes L + las 11 `<f>` + inventario de hojas + base interna `N3`), nunca edición silenciosa de literales. Drift gate: `GeometriaPlantillaEstandarizadaT0Tests`.

**Regla de proceso:** regen SIEMPRE a ruta temp FRESCA con el OFICIAL como única entrada y rutas explícitas (prohibido `Docs/Insumos` raíz en lo nuevo). El oficial se verifica intacto por SHA256 (hash T3: `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).

## Doctrina W-1 evalúa + writer sella (Plan 40 — vigente)

**Principio rector (`los cachés de fórmulas NO son oráculo; el gate EVALÚA y el writer SELLA`):** el gate W-1 (`ValidadorBasico` + `ValidacionOracleReader`) deja de leer los cachés `<v>` heredados de la plantilla y **evalúa** los 5 sub-bloques booleanos de `VALIDACION_TOTAL` desde los valores realmente escritos en la salida; antes de evaluar **aserta el TEXTO `<f>` esperado** (fixture T0) contra el workbook (si la plantilla cambia la fórmula, el gate falla ruidosamente con `ERR-PLANTILLA` en vez de verificar la fórmula equivocada). El writer **sella** los `<v>` computables por **read-back de lo escrito** (solo `<v>`, **jamás `<f>`**, **jamás desde la intención del dominio**), en un camino separado `SellarCacheFormula` con assert de texto-`<f>` + preservación byte de `<f>`; el guard anti-fórmula de `EscribirValorNumerico` **NO se relaja**. Sellado corre sobre la copia/salida, nunca sobre el oficial (`ValidarNoInPlace` intacto).

**Criterio de éxito ESTRICTO (Enmienda E-1 — ratificado por el Ingeniero):**

- **E1-1:** **prohibido `SeGatea=false` como forma de pasar.** El evaluador resuelve la cadena COMPLETA de los 5 sub-bloques desde lo escrito (incluido el espejo de remuneración / `Reporte Componentes R1`) y emite TRUE/FALSE por celda; "no evaluable" solo con demostración numérica + cadena citada, y aun así queda como **hallazgo**, no como verde.
- **E1-2:** un sub-bloque que no cierra es un **DEFECTO REAL** con **frente propio**, nunca un gate a apagar/re-etiquetar/sellar.
- **E1-3:** T4 (regresión) **no** declara ÉXITO si algún sub-bloque quedó sin evaluar (exige 5 evaluados TRUE, o el defecto E1-2 declarado y escalado).
- **E1-4:** bloqueante **Q1** de la familia F-NEW-1: el oficial referencia `VALIDACION_TOTAL!C3:C8 → 'Recaudo *'!F21:F26` **fijo**, pero en Q1 el recaudo se escribe en columna **D** ⇒ `C9=0≠F10` ⇒ `C15` FALSE en Q1; Phase 0 reescribe el `<f>` por `NumeroQuincena` (dominio, nunca contenido).
- **E1-5:** la tabla de evaluación §2.2 y el veredicto T5 quedan vinculados a E1-1/E1-2: "requiere decisión" **no** es salida válida para los 5 sub-bloques de W-1 (sí para el resto de la clase que T5 barra).

**Tabla de veredictos de la clase (T5, resumen):**

| Celda | Veredicto | Detalle |
|---|---|---|
| `VALIDACION_TOTAL!C15` (`C14=C9`) | **se evalúa** | cadena de recaudo 100 % escrita por la app; `C15 = (F10 = Σ C3:C8)` |
| `VALIDACION_TOTAL!D25` (`D24=C9`) | **se evalúa** | recaudo |
| `VALIDACION_TOTAL!O25` (`O24=N9`) | **se evalúa** (cadena de remuneración) | depende del espejo R1 → hoy FALSE por E1-2 (frente abierto) |
| `VALIDACION_TOTAL!D34` (`D33=D109`) | **se evalúa** (cadena de remuneración) | idem E1-2 |
| `VALIDACION_TOTAL!F34` (`F33=D109+N109`) | **se evalúa** (cadena de remuneración) | idem E1-2 |
| Cadena recaudo `C3:C8→C9→C14→C15` | **se sella** (read-back) | allow-list del writer, values-only |
| F-NEW-1 (`DetRetri{per}!L32:L36`/`O32:O36`, `DetValiRetri{per}!I24:I29`, `INTERVENTORIA!S26:S31`) | **requiere decisión** | auto-chequeos con columna `M` fija; su T0 propio (ver abajo) |
| Resto de la clase (cachés viejos heredados: `O3:P9`, `VALIDACION_*`, `DetValiRetri` D24:D29, `DetRetri` D32:D36) | **invariante** | falsos-verdes latentes documentados; no se sellan sin read-back |

**Regla F-NEW-1 + bloqueante Q1 (Plan 40):** el oficial trae fórmulas **fijas** que deben variar por quincena y no varían: el recaudo `'Recaudo *'!{D,F}21:{D,F}26` (Q1 escribe en **D**, Q2 en **F**) y los auto-chequeos con `INTERVENTORIA!M26:M31` (Q1 usa `N26:N30`). **Phase 0 las reescribe por `Periodo.NumeroQuincena`** (dominio, **nunca** por contenido). **Allow-list Phase 0 vigente: 68 celdas** = 16 del Plan 39 (11 tokens de período + 5 `L9:L13`) **+ 36 de recaudo** (Q1/Q2 de las 5 hojas `Recaudo *` × filas 21:26 en columnas D/F) **+ 16 de F-NEW-1**; `ValidadorBasePeriodo` exige la **variante del run** y rechaza una tercera forma con `ERR-PLANTILLA`. Prohibido sellar TRUE donde la divergencia es real (mentir en caché es peor que el FALSE honesto).

**DEFECTO ABIERTO E1-2 y su frente (espejo R1 — Plan 41):** el espejo R1 **no puebla** correctamente el sub-bloque por empresa del **ASE5–ENEL** (`Reporte Componentes R1!F633/F663/L633`, slots de valor plano en el oficial) ni el de **ENEL ASE3** (`F345`), **OCCIDENTE ASE1/ASE3** y **Reciprocidad ASE2/ASE4** → las anclas `REMUNERACION_*!D9:D13` quedan cortas/cero (`REMUNERACION_ENEL!D13 = 0` vs `11.701.157.844,90` del manual; `ENEL ASE3 = 1.307.680,00` vs `13.997.194.578,87`; etc.) → `O25`/`D34`/`F34` FALSE en Q1 y Q2 → la app **falla honestamente** (15 errores, sin `ERR-VALIDACION` silencioso) hasta que el frente lo arregle. **Causa raíz:** `EscritorValoresSlotR1` agrupa los slots por firma A/B/C/D/E y los consume con un **cursor físico** (`GrupoFirma.SiguienteLibre`), sin respetar la **sección** (frontera inmediatamente arriba) de la fila fuente; cuando la fuente trae menos filas que slots fijos, las filas caen en los primeros slots y las anclas por empresa quedan vacías. **Diagnóstico completo:** `Docs/T0-Plan41-Frente-EspejoR1-Ase5Enel.md` (tabla ancla→esperado→observado por empresa×ASE, veredicto lectura-vs-escritura, causa raíz citada, opciones de fix). **Segundo hueco (frente distinto):** `AJUSTES - SF-T!D47:D51` = 0 porque la app desplaza +1 las filas de `SALDOS POR NOTA` (fila `Vlr Intereses` extra en el bloque oficial). **`INTERVENTORIA!S26:S31`** (`Q26=R26`, `R=J15:J19` anual sin variante `N`) requiere **T0 propio de la tabla J** (Q1→`N26:N30`); no es invariante.

**Trampa de comparación (R-TRAMPA-ANCLA):** las anclas de fila **difieren** entre el manual legacy y la geometría oficial (p. ej. ENEL: manual `F63/F216/F353/F478/F568` ↔ oficial `F85/F264/F425/F578/F690`; el bloque ASE5 empieza en 501 legacy vs 603 oficial). **Nunca** compares direcciones entre plantillas distintas; la correspondencia es por (empresa, ASE) y el valor esperado es el del manual en SU ancla. En el oficial coexisten dos familias: `CONSOLIDADO_TOTAL RECAUDO!D9:D13` → `F75/F254/F415/F568/F680` (TOT_OPT por ASE) y `REMUNERACION_*!D9:D13` → `F80..F705` (subtotal por empresa).

## Doctrina colocación por SECCIÓN del espejo R1 (Plan 41 + Enmienda E-1 — vigente)

**Principio rector (`la fila fuente va al slot de su SECCIÓN resuelta, y punto`):** el escritor values-only por slot (`Remuneracion.Infrastructure.Excel.EscritorValoresSlotR1`) deja de colocar por **cursor físico** (ni global, ni por firma, ni pool compartido — Enmienda E-1, **E1-1**) y coloca cada fila fuente en el slot cuya **SECCIÓN** (frontera inmediatamente **ARRIBA**) coincide. La sección la resuelve el predicado puro de Core `SeccionEspejoR1.PorFronteraArriba(filas, i)` (Plan 41/T1): `Mes`→Oportuno, `Aplicacion`→Aplicación, sin frontera arriba→1.ª sección; el `Ordinal` distingue ocurrencias (Aplic1/Aplic2, Oportuno1/Oportuno2). `R1FirmaInterior.EsAplicacionPorFrontera` (frontera por DEBAJO, Plan 35-T2) queda **INTACTO** (otro consumidor, otra semántica). El writer agrupa los slots por **(firma + sección)** (`ConstruirGruposSeccion`) y consume dentro de la sección en orden; `LeerSlots`/`RangoDatosAse` (geometría) intactos.

**El catálogo de sinónimos se consume para resolver la sección (Enmienda E-1, E1-2):** el match de rótulo de empresa pasa por `SinonimosEmpresaR1.SonMismaEmpresa` (Plan 34, catálogo abierto; p. ej. fuente `RECIPROCIDAD` ↔ slot `NUEVO ESQUEMA`) **antes** de resolver la sección. Cero ramas por nombre de empresa ni por período (D-H).

**Fail-fast accionable (E1-3):** una fila cuya (firma, sección) no tiene slot en el contrato T0 **NO se coloca por posición** → el run FALLA con `ERR-SIN-SLOT` nombrando **ASE + firma A|B|C|D|E + rótulo de empresa + sección esperada + ocurrencia**. Prohibido resolver por proximidad, orden físico o cualquier heurística posicional. Test del fallo: `SeccionEspejoR1Tests.T1_FilaSinSeccionResoluble_ErrSinSlotConRotuloYSeccion` + `T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback`; test de sinónimo por catálogo: `T1_FilaSinonima_ResuelvePorCatalogoEnSuSeccion`.

**Cierre Q2 (Plan 41/T1):** los 6 pares empresa×ASE en rojo por E1-2 cierran al valor del manual ±0.5 por sección (Reciprocidad ASE2 `11.411.240,00` / ASE4 `5.616.850,00`; ENEL ASE3 `13.997.194.578,87` / ASE5 `11.701.157.844,90`; OCCIDENTE ASE1 `779.194.409,36` / ASE3 `1.229.014.628,26`) y las identidades que ya cerraban siguen verdes (ENEL ASE1/2/4, OCCIDENTE ASE2/4, ENERBIT ASE5 `26.864.220,00`). Tests: `SeccionEspejoR1Tests` (`T1_CierraParEmpresaAse_AlValorDelManual`, `T1_IdentidadParesQueCierran_AlValorDelManual`).

**OCCIDENTE ASE5 — DECLARADO (Enmienda E-1, E1-5):** divergencia de **ESTRUCTURA** de fórmula, no de colocación. El oficial `F700 = F647+F673+F612-L612-L647` (3 términos) es internamente consistente con sus ASE1–4; el manual `F578 = F530+F551-L530` (2 términos) omite el término de la 1.ª sección. Con la fuente julio-2, Δ = **417.519,62** = la fila de la 1.ª sección de OCCIDENTE. El par queda **DECLARADO** (nunca verde, nunca forzado al manual): la pregunta va a la **UAESP** (¿la remuneración de OCCIDENTE ASE5 incluye la 1.ª sección?). Diagnóstico: `Docs/T0-Plan41-MiniT0-OccidenteAse5-SaldosNota-TablaJ.md` §1.

**DEFECTO ABIERTO (Enmienda E-1, E1-4/E1-6) — fuentes Q1 incompatibles con la geometría de secciones del oficial:** al quitar el fallback, la corrida Q1 (`2026071`) FALLA con `ERR-SIN-SLOT` porque las fuentes de **1.ª quincena** (`Docs/Prueba Julio-1/Insumos/*/Recaudoporcomponente_*.xlsx`) **no traen las filas-frontera `Aplicacion nuevos x reversion`** (su estructura es `Primera / Oportuno#1 / Oportuno#2`, sin Aplicación), mientras el oficial declara 5 secciones fijas (`Primera / Oportuno#1 / Aplic#1 / Oportuno#2 / Aplic#2`). Casos medidos: **ASE2** trae `NUEVO ESQUEMA` en su 1.ª sección y el oficial NO declara ese slot (su 1.ª sección `F110:F119` solo tiene `ENEL F112` / `OCCIDENTE F119`); **ASE5** trae `ENERBIT` en Oportuno#1/Oportuno#2 y el oficial declara `ENERBIT` SOLO en Aplic#1 (`F637`) / Aplic#2 (`F666`). El catálogo de sinónimos NO crea slots (solo iguala rótulos). El único resolvedor es el **fallback posicional PROHIBIDO** (E1-1) → requiere **contrato nuevo** (modelo de sección Q1-aware / alineación fuente↔oficial) → **STOP y reporte para plan hijo** (E1-6). Impacto medido: 8 tests Q1 (e2e `RegresionOficialPlan39Tests`, `BasePeriodoPreflightTests.S5`, `ProcesadorPeriodoTests`, `RegresionEvaluaSellaPlan40Tests.T4`, `IntegridadCalcChainYSelloFechasTests`, `GoldenValidacionesCruzadasTests.CapaA_Q1`, `AjustesSfTTests`, `MicroFixesHu17Tests.S3`) pasan de verde a rojo por esta causa — **NO se reintroduce el fallback** (E1-4). El **residuo Q1** de `E1-6` (`8.392.943.748,41`) es la MISMA clase (colocación del espejo R1): medido con el fallback (solo diagnóstico), es la divergencia del sub-bloque `D34` (`49.817.151.072,09` espejo vs `58.210.094.820,50` `CONSOLIDADO!D109`); su resolución exige el contrato nuevo.

**Cierre Q2 medido (T4-B):** la corrida Q2 (`2026072`) desde el OFICIAL produce la salida; el evaluador W-1 termina en `ERR-VALIDACION` con `O25`/`D34`/`F34` FALSE en los 5 ASE y **Δ = 417.519,62 exacto** (el par OCCIDENTE ASE5 declarado). `C15`/`D25` TRUE. Es decir: la divergencia Q2 del espejo ya es SOLO el par declarado; `CONSOLIDADO_TOTAL RECAUDO!D9:D13`, `DetRetri`/`DetValiRetri` y las empresas sin divergencia intactos (R-INTACTO).

## Doctrina divergencia declarada con reconocimiento explícito (Plan 42 — vigente)

**Principio rector (`producir y cerrar sin silenciar; la declarada es la ÚLTIMA puerta`):** la app **produce el archivo y cierra el run** cuando la ÚNICA divergencia que queda es **DECLARADA** (externa, con evidencia citada y dueño fuera de la app), **exigiendo reconocimiento explícito del usuario por corrida** y dejando **DOS huellas obligatorias** (reporte durable + log). Una divergencia **NO declarada sigue fallando en duro** con el criterio estricto E1-1..E1-5 intacto (Plan 40). Nada de "avisar y seguir": declarada ≠ tolerada en silencio.

### Divergencia declarada vs NO declarada (taxonomía T0 — `Docs/T0-Plan42-Taxonomia-DivergenciasDeclaradas.md`)

- **Declarada:** diferencia **EXTERNA por construcción** (discrepancia oficial-vs-manual con dueño fuera de la app). Exige **6 campos citados** (celda(s), origen del monto, dueño, cita de evidencia, fecha y condición de retiro); sin ellos la entrada **NO aplica** (nunca un verde silencioso). Vive en un **catálogo de DATOS** (`Remuneracion.Core.Models.DivergenciasDeclaradasCatalogo`, patrón Plan 34): tabla citada + matcher puro `Buscar(celda, observado, esperado)`, **cero ramas por empresa/período** (D-H). Añadir una entrada futura = añadir una FILA citada, nunca lógica nueva.
- **NO declarada:** todo **defecto interno** (lectura, cálculo, colocación, sellado, naming, preflight, caché). **NUNCA es declarable**; sigue el criterio estricto con el mismo nombre accionable `ASE n · VALIDACION_TOTAL · <celda>`, mismo fail-fast y misma severidad.
- **Catálogo vigente — primera (y hoy única) entrada:** `OCCIDENTE-ASE5` — celdas `O25`/`D34`/`F34`, **Δ = 417.519,62**, **dueño UAESP**, cita `Docs/T0-Plan41-MiniT0-OccidenteAse5-SaldosNota-TablaJ.md` §1, **condición de retiro**: "se retira cuando la UAESP resuelva (se elimina la fila, no se adapta el código)".

### Clasificación en Core (una sola decisión; UI/CLI solo presentan)

`ValidadorBasico.ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots)` reutiliza `Validar(...)` (criterio estricto intacto) y clasifica cada sub-bloque FALSE: si corresponde a una entrada vigente del catálogo → `Declarada` (cita la entrada); si no → FAIL duro. El veredicto por celda + referencia a entrada + montos/Δ viaja **aditivo** en el resultado (`ClasificacionDivergenciasDeclaradas` en `ResultadoProcesoPeriodo`; vacío = comportamiento anterior, compat HU-13). **D-F (orden vinculante):** primero todo lo interno; la declarada es la **última puerta** — si hay CUALQUIER no declarada (`HayNoDeclaradas`), el run FALLA en duro aunque la declarada esté reconocida.

### Doble huella obligatoria (D-E) — sin reporte NO hay ÉXITO (fail-closed)

Al reconocer, `ProcesadorPeriodo` emite:
1. **Reporte durable** junto a la salida: `<archivo de salida sin extensión> - Divergencias declaradas.txt` (p. ej. `Remuneración 202607-2 Total - Divergencias declaradas.txt`), UTF-8 sin BOM, texto de negocio (celda, observado vs esperado, Δ, dueño, evidencia, fecha, condición de retiro, qué se reconoció). Implementación: `ReporteDivergenciasDeclaradasWriter` (Infrastructure) tras `IReporteDivergenciasDeclaradas`; sin emisor configurado → `ERR-INESPERADO` (nunca cierra sin reporte).
2. **Huella en el log** (Serilog): entrada, celda, Δ, `RunId` y la aceptación (`RegistrarHuellaDeclaradas`).

### Reconocimiento por corrida (D-D, anti-costumbre) — UI y CLI

- **UI WinForms (dos pasadas):** 1.ª pasada SIN aceptaciones → si lo único pendiente es declarada, el procesador LANZA con la clasificación; `Form1` presenta `FormReconocimientoDivergencia` (acto consciente, nada pre-marcado, nada persistente) y, con el acto, **re-ejecuta** con los ids reconocidos. Sin acto no hay cierre (`ERR-VALIDACION`). El emisor del reporte viaja en el **9.º argumento de `ProcesadorPeriodo`** (`IReporteDivergenciasDeclaradas`); los ids reconocidos viajan en `SolicitudProcesoPeriodo.DivergenciasDeclaradasAceptadas`.
- **CLI no-interactivo (D-G):** flag `--aceptar-divergencia-declarada <id>` (repetible, cita UNA entrada, no "todas"). Sin flag la declarada bloquea (`Validacion` = 1).
- **Código de salida 6 = `CodigosSalida.DivergenciaDeclaradaReconocida`** (DISTINTO de `Ok = 0`): el orquestador del lote distingue un cierre limpio de uno con deuda externa reconocida. Contrato 0..5 intacto.

### Abierto con dueño (cierre del Plan 42)

- **Q1 (2026071) sigue DECLARADO:** su bloqueo es un **contrato de sección faltante** (fuentes de 1.ª quincena sin filas-frontera `Aplicacion nuevos x reversion`) → **plan hijo** con su T0 de bytes (E2-5). Los **8 rojos de Q1** quedan declarados con esa causa: ni verdes por apaño, ni tocados, ni silenciados.
- **OCCIDENTE ASE5 pendiente de la UAESP:** el veredicto de negocio (¿la remuneración de OCCIDENTE ASE5 incluye la 1.ª sección?) lo arbitra la UAESP; al resolver, la entrada se RETIRA por su condición de retiro (no se adapta el código).
- **Pruebas manuales del Ingeniero:** las causales (sin reconocimiento → bloqueo; con reconocimiento → cierre; mixta → FAIL duro) se prueban usando el **reporte durable** como evidencia de qué se aceptó y de quién es la deuda.

**Evidencia de cierre (T5, 2026-10-14):** build 0 warnings; suite **39 Failed / 873 Passed / 912 Total** con **set-diff vacío** contra la línea base de entrada (mismos 39, 8 de Q1 declarados); e2e Q2 desde el OFICIAL en ambas direcciones (sin reconocimiento → exit 1 con mensaje accionable; con `--aceptar-divergencia-declarada OCCIDENTE-ASE5` → exit 6 + archivo + reporte + log); SHA256 del oficial intacto `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`.

