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
- **LA PLANTILLA ESTÁNDAR ES INTOCABLE (regla absoluta, ratificada 2026-10-22):** **NUNCA** modificar fórmulas, **NUNCA** agregar ni eliminar filas, **NUNCA** agregar ni eliminar columnas. El pipeline es **values-only** sobre celdas existentes; el guard anti-fórmula (`ERR-PLANTILLA`) no se relaja. Las únicas excepciones declaradas con allow-list + gates son la Phase 0 offline (Plan 39/40) y el espejo R1 5→1 (Plan 21); nada más toca `<f>` ni la geometría.
- **ASE4 (Bogotá Limpia)**: sin columna "Especiales" — chequear headers dinámicamente; si ausente usar 0.
- **SALDOS POR NOTA** y **RETRIBUCION NEGATIVA**: solo quincena 2 (`Periodo.NumeroQuincena == 2`); `AjustesSfT` = 0 en quincena 1.
- **Tolerancia:** diferencias ≤ ±0.5 entre calculado y fuente son aceptables (redondeo).
- **Filas dinámicas:** buscar por headers, NO por número de fila.
- **R4:** valores de reversión negativos en columna 4.
- **Redondear a entero** antes de escribir en CONSOLIDADO columna D (`DetRetri`).

## Modelo de la plantilla estándar (doctrina transversal — reunión con la propietaria, 2026-10-21)

**La plantilla estándar define la ESTRUCTURA; la fuente del período solo aporta los VALORES.**

| Elemento | Dónde vive (ej. `Reporte Componentes R1`) | Qué es |
|---|---|---|
| **Componente** | Encabezado de columna: `Componente TDF`, `Componente TTL`, `Componente TVIAT`, `Aprovechamiento`, `CCSA Prest.Aprov.`, `SERVICIO ESPECIALES`, … | Concepto económico del valor |
| **Fuente de recaudo** | Columna **C**: `ENEL`, `OCCIDENTE`, `NUEVO ESQUEMA`, `ENERBIT`, … | Origen del recaudo |
| **Canal de recaudo** | Columna **D**: `H`, `O`, `T`, `E`, `1`, `2`, `5`, … | Canal dentro de la fuente |
| **Concepto** | Columna **E**: `Vlr Servicio`, `Vlr Intereses`, `Total` | Tipo de valor del canal |

Jerarquía: cada **fuente** tiene sus **canales**; cada canal, sus **conceptos**; los valores son los **componentes**. La plantilla define qué fuentes/canales/conceptos existen por cuadro (no todos los cuadros traen los mismos) y **son los que se deben mapear**.

**Columnas (verificado en bytes, obs Engram #2124):** la plantilla estándar trae los componentes en **F..Z** (21 columnas; incluye `Exonera-*`, que la fuente NO trae, y `Z = Rec.Intereses`); la fuente trae **F..U** (16). Los componentes se mapean **POR ENCABEZADO**, nunca por posición (ej.: `Rec.Intereses` fuente `U` → estándar `Z`); un componente sin contraparte en la fuente queda en **0**. La plantilla contiene **todas las filas posibles**; el llenado es dinámico desde la fuente.

**Brecha conocida (2026-10-16, obs #2124):** la App interpretaba C/D/E de raíz y escribía solo `F` (Total) y `L` (SERVICIO ESPECIALES), dejando componentes en 0. Los planes 41-49 avanzaron el mapeo por firma/canal/encabezado; la **opcionalidad de la fuente** y la **unificación Q1/Q2** quedan pendientes de plan (ver rojo 2026101 ASE2).

**Reglas operativas (ratificadas por el Ingeniero 2026-10-21):**
1. **La plantilla manda en la estructura**; el mapeo es por **ETIQUETA** (fuente/canal/concepto), nunca por posición ni por ocurrencia.
2. **Lo que la fuente no trae no se exige**: la celda destino queda en **0** (la plantilla ya está en 0). Es el mismo criterio de la plantilla estandarizada: se coloca lo que viene.
3. **Q1 y Q2 se tratan igual** en este manejo: no hay reglas particulares por quincena.
4. **Nunca abortar por ausencia de una fuente de recaudo**; registrar la ausencia y continuar.
5. Los **invariantes de cierre de la fuente** (`A='Componente' B='Total'`, `A='Subs/Cont' B='Total'`, `A='Total'`) siguen siendo fail-fast si faltan (T0e, Plan 21): son estructura del reporte, no una fuente opcional.
6. **Las fuentes Y los canales pueden venir o no venir** (indicación de la propietaria, ratificada 2026-10-22): una sección o canal ausente **no es estructura rota**; su celda destino queda en **0**. Es normal que una fuente no traiga una sección completa (p. ej. `AFaseo`/primera sección). Prohibido resolver por ocurrencia/índice: una sección ausente **no debe duplicar** el bloque de otra sección (`TotalMenor` = 0, nunca reutilizar `filas[count-1]`).
7. **La señal de sección es la frontera** (`Mes`, `Aplicacion`, `Componente`), no el índice. El mapeo por empresa debe resolver por **(firma + sección + canal)** como el espejo R1 (Planes 45/46), nunca por posición.

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

## Doctrina Roles R1-Q2 por firma sobre la secuencia espejo (Plan 25 — vigente; dispatch por quincena SUPERSEDIDO por el Plan 54)

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

## Doctrina contrato de sección PERIOD-FREE del espejo R1 (Plan 45 — SUPERSEDIDA EN PARTE POR EL PLAN 46)

> **SUPERSESIÓN (2026-10-17):** el tipado por frontera ARRIBA y el cero documentado de `Vlr Intereses`
> descritos en esta sección quedan **SUPERSEDIDOS** por la doctrina del Plan 46 (al final de este
> archivo): sección por frontera de **CIERRE** + **canal efectivo**, y `CerosDocumentadosR1` sin
> patrones. Lo demás (Alinear monotónico, `ERR-SIN-SLOT`, neutralidad Q1 plan-then-write, frontera
> arriba intacta para Plan 35-T2) sigue vigente.

**Principio rector (`la sección de la fuente se ALINEA con la del oficial por la FORMA, no por el tipo/ordinal`):** el escritor values-only por slot (`EscritorValoresSlotR1`) deja de dejar en `null` la sección de las filas sin empresa en C; ahora TODA fila participante recibe su **sección-contrato**, resuelta por el predicado puro de Core **`ContratoSeccionPeriodFreeR1.Alinear`**: emparejamiento MONOTÓNICO por tipo (la k-ésima sección fuente → la próxima sección oficial del mismo tipo aún no consumida). El oficial declara siempre 6 secciones `[1.ª, Oportuno#1, Aplic#1, Oportuno#2, Aplic#2, Oportuno#3]`; la fuente del período define su forma.

**Formas absorbidas sin ramas (3 períodos probados, `Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md`):**
- `[P,O,A,O,A,O]` (mayoría) → identidad;
- `[P,A,O,A,O]` (ASE5-julio2 invierte: arranca con `Aplicacion`, 5 secciones; también ASE1-Q1) → `[1.ª, A#1, Oportuno#2, A#2, Oportuno#3]` (salta el Oportuno#1 oficial): el bloque `Componente/Subsidio` fuente cae en el **Oportuno#2** oficial (649-659), no en el Oportuno#1;
- `[P,O,O,O]` (agosto-ASE3 / Q1-ASE2/ASE5, sin `Aplicacion`) → `[1.ª, Oportuno#1, Oportuno#2, Oportuno#3]` (salta las Aplic del oficial).

Las **filas-frontera** (`Mes/Total`, `Aplicacion…/Total`) viajan con la sección que ABREN (su sección-contrato es la de la sección que abren, no la de arriba): el `Mes/Total` de ASE5-julio2 que abre el `Oportuno#1` fuente va a **F648** (`Oportuno#2`), y el que abre `Oportuno#2` → **F674**; el `Oportuno#1` oficial (**F613**) queda en cero. Las filas **AFaseo** quedan `null` (se localizan por firma, dual-firma intacta).

**Cero documentado (Plan 45 §2.2 — `CerosDocumentadosR1`):** una fila fuente sin slot cuya `(firma, sección-contrato)` está en la lista T0 NO es error: queda en el cero del template (no se escribe, no se simula) y se reporta al dueño. Hoy: la fila `Vlr Intereses` de nivel canal (A/B/C/D vacíos) en la **1.ª sección** — el **hueco de PLANTILLA** (ASE5-julio2 R8/R11: `O` 402,40 / `T` 7.697,23; el manual legacy los registraba; la 1.ª oficial no trae `Vlr Intereses` en NINGÚN ASE). Los Totales ya llevan el dinero (`O` 50.323,46 = 49.921,06 + 402,40; `T` 314.490 = 306.792,77 + 7.697,23): solo se difiere el desglose. **NO se toca la plantilla** (SHA256 intacto); es un reporte al dueño.

**`ERR-SIN-SLOT` residual:** fail-fast enriquecido (ASE + firma A|B|C|D|E + rótulo de empresa + canal + concepto + sección esperada + ocurrencia) SOLO para fila con `(firma, sección-contrato)` sin slot FUERA de la lista de cero documentado. Prohibido fallback por orden físico/proximidad (Plan 41 E1-1).

**Micro-punto F654/F661 (cerrado):** el duplicado `−187.855,98` nacía en el ESCRITOR (cursor físico de filas sin C), no en el lector (una emisión por fila). La sección-contrato lo tapa: `Subsidio`/`D=H` van a su slot de sección, sin duplicar (fixture `ContratoSeccionPeriodFreeH1Tests`).

**H1.4/H1.5 (review F1 plegadas):** encabezado de fuente sin columna destino en el template → `Log.Warning` accionable (ASE + fila + encabezado) + `EncabezadosSinDestino` en el resultado (nunca descarte silencioso; hoy inocuo: F..U mapea 1:1 a F..Z). `BloqueEspejoAseInputs.TieneColumnaEspeciales` queda como dato FIXTURE/diagnóstico (sin consumidor en producción tras el llenado por encabezado; semántica intacta).

**RESTRICCIÓN DURA Q1 (vigente):** las formas Q1 se **CITAN** en el T0 (Q1-ASE2 `D=5`/`NUEVO ESQUEMA` en la 1.ª sin slot; Q1 sin `Aplicacion`), pero Q1 **NO se resuelve** aquí; cualquier cambio de comportamiento Q1 exige autorización explícita separada. El T0 de Q1 solo LEE bytes (read-only).

**Neutralidad Q1 (decisión del Ingeniero 2026-10-16, opción 1 — implementada):** el escritor aplica la **sección-contrato SOLO si el bloque ENTERO es colocable con ella** (plan-then-write, `TryPlanificar`). Si no lo es (la forma del período no está cubierta por el contrato — p. ej. Q1-ASE2 trae el canal `D=5` en la 1.ª, que el oficial no declara), reproduce el **comportamiento HISTÓRICO del Plan 41** para ese bloque (veto de sección solo en filas-dato de empresa; cursor por firma). Así Q1 conserva **byte-idéntico** su primer `ERR-SIN-SLOT` (`NUEVO ESQUEMA`, no `D=5`) y **ningún test ni geometría Q1 cambia**; set-diff de suite **vacío** contra la línea base de 39. La decisión es por **forma colocable** (geometría pura), no por período/empresa (D-H).

**Gate de forma (re-congelamiento):** un período con forma nueva NO se parchea con ramas; la regla es geometría pura y una forma inédita exige un **T0 nuevo** (`ContratoSeccionPeriodFreeT0Tests`), nunca edición silenciosa.

## Doctrina sección por CIERRE + canal efectivo del espejo R1 (Plan 46 — vigente, supersede al 45)

**Principio rector (`la sección se tipa por la frontera que la CIERRA; la fila sin D se coloca por su canal efectivo`):** la corrida real del App del 2026-10-17 (`Docs/Prueba Julio-2/Resultado/8/`) probó, celda a celda, 2 defectos del contrato del Plan 45. El Plan 46
(`plans/46 - correccion-mapeo-interior-r1-canal-efectivo-y-seccion-por-cierre.md`) los cierra con dos predicados puros de Core:

1. **Sección por CIERRE** (`ContratoSeccionPeriodFreeR1.SeccionPorCierre`, §2.2): cada sección (fuente y oficial) se tipa por la frontera **inmediatamente debajo** que la cierra (`Mes`→Oportuno, `Aplicacion`→Aplicación, ordinal = ocurrencia de esa frontera); las filas-frontera se tipan por la sección que **ellas cierran**; la cola sin frontera debajo se queda con la de arriba; AFaseo → `null`. La 1.ª sección **deja de ser `Primera`**: si cierra con `Aplicacion`, es `Aplicacion`.
2. **Canal efectivo** (`ContratoSeccionPeriodFreeR1.CanalEfectivo`, §2.1): para una fila **sin `D` propio** y sin `A`/`B` (`Vlr Servicio`, `Vlr Intereses`, `Subsidio…`), su canal = el `D` de la próxima fila `…/Total` hacia abajo con `D` no vacío y ≠ literal `Total`. Mismo cálculo en slots del oficial. La llave de colocación pasa a ser **(firma + sección + canal)**.

**Formas por cierre y mapeo (verificado en bytes; fixture `ContratoSeccionPeriodFreeT0Plan46Tests`, 26 verdes):**
- `[O#1,A#1,O#2,A#2,O#3]` (julio2 ASE1-4, agosto mayoría, Q1-ASE3/4) → identidad.
- `[A#1,O#1,A#2,O#2]` (julio2 ASE5, Q1-ASE1) → `[A#1(615-628), O#2(630-647), A#2(649-659), O#3(661-673)]`; la 1.ª oficial (603-613) y el AFaseo (614) quedan en cero. `F612`=0; OCCIDENTE de la 1.ª (R13) → **F628**; `Vlr Intereses` 402,40/7.697,23 → **F623/F626**.
- `[O#1,O#2,O#3]` (agosto-ASE3, Q1-ASE2/5) → `[O#1,O#2,O#3]`; **no colocable** por el contrato (la fila `Componente` fuente cierra con `Mes` y la del oficial con `Aplicacion`) → el plan-then-write reproduce el comportamiento histórico del Plan 41 (neutralidad Q1).
- **Frontera arriba INTACTA:** `SeccionEspejoR1.PorFronteraArriba` y `R1FirmaInterior.EsAplicacionPorFrontera` (consumidor Plan 35-T2) no cambian de API ni de semántica.

**Cero documentado RETIRADO (FP-3):** `CerosDocumentadosR1` queda con lista de patrones **vacía**; el mecanismo se conserva. Las 2 entradas de `Vlr Intereses` de la 1.ª de ASE5-julio2 se retiran (tienen slot real). El texto §9 del T0-Plan45 (reporte al dueño del "hueco") queda **WITHDRAW**.

**Invariantes intactos:** values-only; cero `InsertRow`/`DeleteRow`; cero `<f>` nuevo; `R1FirmaInterior.EsDatoEmpresa` (Plan 41) intacto; cero ramas por período/empresa (D-H); Q1 intocado (set-diff de Q1 vacío); `DivergenciasDeclaradasCatalogo` **intacto** (no se decide Vía A/B aquí). La auditoría de la propietaria sigue PENDIENTE (no se declara paridad).

**Efecto medido en la corrida 8→9 (julio-2, 5 ASE):** los cambios de celda son SOLO movimientos de fila-fuente, 100 % trazados por DEFECTO 1 (canal) o DEFECTO 2 (sección). Los 24 movimientos son: ASE1 R8 (F9→F11); ASE2 R8, R17-R18, R20-R21, R54, R56; ASE3 R5 (F289→F294); ASE4 R8, R20-R21, R54; ASE5 R5-R13 (F603-F612 → F619-F628, con R8/R11 de 0 a F623/F626) y R35/R37 (F650/F653 → F653/F655). La corrida Q2 **cierra OK (código 0) sin divergencia declarada**: `VALIDACION_TOTAL!O25/D34/F34` pasan de FALSE (Δ 417.519,62) a TRUE (Δ 0); **pendiente de decisión del Ingeniero/UAESP** re-basar el catálogo y sus 11 tests (el Plan 46 prohíbe tocarlos).

## Doctrina detalle R2/R4 por `(firma + canal efectivo compuesto)` + ceros documentados + `Otros` (Plan 47 — vigente)

**Principio rector (`el desglose-detalle de R2/R4 se coloca por firma + canal efectivo compuesto; las filas sin slot están documentadas o fallan`):** el T0 del Plan 47 (`Docs/T0-Plan47-Revision-Pestanas-LenteR1.md`) probó por bytes que el mapeo LCS por firma casaba **por orden** y producía corrimientos (gemelo D1 del 46), falsos huecos y omisiones silenciosas. El H1 (autorizado) lo corrige con un predicado puro de Core.

1. **Canal efectivo del detalle** (`CanalEfectivoDetalleR2R4.CanalEfectivo`, Core puro): la fila se ubica por su **firma + el rótulo del grupo que la CIERRA** (la primera `…/Total` hacia abajo con rótulo de grupo). La llave de colocación es `(firma + canal efectivo)`.
   - **Identidad de la fila:** las filas con etiqueta propia **fuera** del grupo (roles `Componente`/`Subs/Cont`, empresa, `Total` del bloque) se localizan por firma y no participan del canal. En R2 eso exige distinguir "etiqueta propia" (`A|B|C`) de "etiqueta de grupo" (solo `C`): la fila de sección `A=Componente B=Total` NO toma canal.
   - **Canal COMPUESTO (segundo cierre, W-3):** en R2 el rótulo se compone `canal/empresa` con la empresa del **primer cierre de empresa por debajo** (`B=ENEL C=Total`) → `E/ENEL`, `E/ENERBIT`, `9/NUEVO ESQUEMA`, `A/EAB`. Así los DOS grupos del mismo canal (R2 ASE5: ENEL y ENERBIT) quedan desambiguados por GEOMETRÍA, no por orden: la corrida real es la misma cifra, pero invertir el orden de la fuente reparte cada grupo en su slot (test `DetalleR2_Ase5_DobleGrupoE_OrdenInvertido_CadaGrupoRecibeLoSuyo`). R4 **no** compone: su rótulo de grupo (columna `A`) YA es la empresa.
2. **Columnas por ENCABEZADO** (nunca por índice): R2 `E..R` (14; `Especiales K` = 0 explícito si la fuente no la trae), R4 `D..O` + **`Otros → R`** (13: falso hueco de columna de ASE4 corregido; `SERVICIO ESPECIALES P` y `Ajuste Recargo Aseo Q` no tienen contraparte en ninguna fuente → 0 de plantilla).
3. **Filas sin slot (E-44.1):** `CerosDocumentadosDetalleR2R4` declara **11 claves / 14 filas** (9 julio-2: ASE2 3 + ASE4 6; 5 agosto-2: ASE2 3 + ASE4 5) con motivo + evidencia citada; el consumo se acota por `Cantidad` **por corrida**. Fuera de la lista ⇒ `ERR-SIN-SLOT` accionable (hoja + ASE + nº de fila + firma + canal). Prohibido el fallback posicional (Plan 41 E1-1).
4. **Fixtures:** `RevisionPestanasLenteR1T0Plan47Tests` (36 verdes: matriz H1 por ASE, identidad de canal en todo par, espejo local vs predicado de producción en julio+agosto, ceros documentados, `Otros`, P3/P5, gate de que ninguna `<f>` consume filas sin etiqueta) y `DetalleR2R4EscrituraTests` (26 verdes, incluye W-2 fail-fast y W-3 anti-orden).
5. **Invariantes:** cero `<f>` nuevo/borrada (13.000 celdas con fórmula byte-idénticas), cero insert/delete, values-only, guard anti-fórmula (`ERR-PLANTILLA`), D-H (cero literales/ramas por período o empresa: los rótulos son DATOS), frente R1 del 46 y `DivergenciasDeclaradasCatalogo` intactos, `SALDOS POR NOTA`/`Recaudo *`/BCE/CONSOLIDADO/AJUSTES/VALIDACION_*/GERENTES_*/REMUNERACION_*/DetRetri* **intactas** (0 `values_changed`). La auditoría de la propietaria sigue PENDIENTE.

## Doctrina adopción plantilla v2 (Plan 48 — vigente)

**Principio rector (`la v2 es la base única desde 2026-10-19; la v1 queda histórica`):** el pipeline consume `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx` (SHA256 `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB`); la v1 (`…C5214`) no la usa el pipeline y se conserva intacta como histórica. **Cambio ratificado 2026-10-21:** el commit `d55e621 "Ajustes nueva plantilla"` reemplazó el binario (`FF5CF7A8…` → `0F192AD2…`) antes del Plan 50 e introdujo 20 rojos de T0 preexistentes (SHA/geometría/fórmulas-caché) y 2 verdes legítimos en `ValidadorTotalesR1GateTests`; los 20 se re-congelan con T0 propio, no a ciegas. El Plan 50 no modificó la v2. **Deuda viva:** `Rem. Anticipos R2!E285/K285` pasaron de literal 0 (v1) a fórmula (v2) → el mapa R2 (1,4) no puede escribirse sin re-congelarse por T0 (`R-MAPA-V2-R2-ASE4`).

- **Geometría R2 v2:** `BloquesR2` = ASE1 `3-33`, ASE2 `67-124`, ASE3 `175-205`, ASE4 `293-350`, ASE5 `404-441`. Cadena de consumidores a `E47/E155/E273/E384/E468` (`WorkbookLeafCellMapQ2`/`PorAse`, `ConsolidadoProtected`, `DetRetri`/`DetValiRetri`). El **cuadro OPORTUNO NO es escribible**: sus filas son fórmula o literal-0; el guard anti-fórmula lanza `ERR-PLANTILLA` si se intenta.
- **Falso hueco cerrado (Plan 47 → Plan 48):** las 9 filas de julio-2 (canal `9` y `Vlr Intereses` del canal 2 de ASE2/ASE4) que el 47 documentó "sin slot" tienen slot en la v2; `CerosDocumentadosDetalleR2R4` conserva **solo 3 claves A/EAB de agosto-2**. Brecha de **53.283.379,41** cerrada (ASE2 `E87/E89=831.919,29`; ASE4 `E313/E314/E315=20.152.773,11 / -2.377,86 / 20.150.395,25`).
- **Control de recaudo (`Valida - Control Recaudo`):** `D7:D10` se pueblan **por etiqueta** de la columna `C` (tabla `ControlRecaudoPorEtiqueta`) desde `Recaudo *!F27`. En la v2 `E7/E10/F10` son fórmulas y el writer **sella su `<v>` evaluado** (`CeldasSelladoControlRecaudo`), sin sobrescribir `<f>`: `F10 = 72.497.949.967` (julio-2) / `77.064.563.298` (agosto-2); `VALIDACION_TOTAL!C14 = F10`, `C15 = 1`.
- **R1 682/702 EXCLUIDO:** las 492 fórmulas del manual (patrón `G→H`, `F682/F692/F702`, shared `H682:AO682`/`H702:AO702`) **no se adoptan** (delta documentado; P5a pendiente). La v2 no las trae. R1 649 solo assert de no regresión (el mirror la escribe).
- **Q1:** fuera de alcance. El end-to-end Q1 usa la **v1 histórica** como dato del test; Q2 corre con v2. **Ratificado 2026-10-19 (P5c): la v2 es la base unica para TODOS los periodos.** Consecuencia obligatoria: procesar una Q1 real con la v2 exige cerrar su frente (hoy falla con `ERR-SIN-SLOT` en R1 ASE2 `||NUEVO ESQUEMA|Total|`, el contrato de seccion faltante de Q1); el end-to-end Q1 se mantiene con la v1 solo como referencia historica hasta esa autorizacion. Cualquier cambio de comportamiento Q1 exige autorización separada.
- **Riesgos residuales declarados:** (a) el fixture `AdopcionPlantillaV2T0Plan48Tests` depende de `%TEMP%\opencode\p47-h2c` (no reproducible en máquina limpia; falla con mensaje accionable); (b) `Insumos.BaseDelPeriodo` deriva **todos** los períodos de la v2 y la propietaria solo auditó julio-2 (**P5c RESUELTA 2026-10-19: el Ingeniero ratifico que la v2 es la referencia para TODOS los periodos y la unica base de aqui en adelante; v1 queda historica. La cobertura de evidencia sigue limitada a julio-2: solo ese periodo fue auditado por la propietaria**); (c) acoplamiento a geometría/rótulos v2: una v3 exigirá re-congelar por T0 con fail-fast visible.

## Doctrina gate absoluto anti-cancelación + regla celda-referenciada (Plan 49 — vigente)

**Principio rector (`una igualdad entre dos derivaciones del MISMO workbook nunca prueba corrección; solo prueba consistencia`):** toda pérdida **común** a ambas derivaciones es invisible por construcción. El gate 2.7 (`VALIDACION_TOTAL!C15/D25/O25/D34/F34`, `ValidadorBasico.ValidarGatesValidacionesCruzadasPorAse`) compara `D33` (lado empresas) contra `D109` (lado ASE) y por eso **canceló** la pérdida de `1.249.810,00` del EXTEMPORÁNEO ASE5 (entraba en ambos lados), mientras el `12.462,26` asimétrico del R2 sí lo rompía. La clase se cierra con un gate **absoluto** donde un operando es el **dominio**, no otra celda del workbook.

### A1-núcleo (obligatorio, Q1 y Q2 sin exención)

- **Hecho:** `Remuneracion.Core.Services.ValidadorBasico.ValidarExtempAbsolutoPorAse`, invocado desde `ValidarGatesValidacionesCruzadasPorAse`; `ClasificarDivergenciasDeclaradas` lo **hereda** sin código extra (reutiliza `Validar`), así que una pérdida A1 es una divergencia **NO declarada** y falla en duro (Plan 42 D-F: la declarada es la última puerta).
- **Evidencia (Infrastructure, read-only):** `ValidacionOracleReader` expone en el snapshot `ExtempResuelto` (= `CONSOLIDADO_TOTAL RECAUDO!D{46+aseId}` resuelto con `ResolutorValorValidacion`: texto-`<f>` + literales, **prohibido `<v>` stale**) y `AnclasExtempR1` (evidencia interior). El reader **no juzga**; el juicio vive en Core.
- **Operando derecho (desviación documentada 2026-10-19):** el plan §2.2 nombra `ConsolidadoAse.Extemp`; en el codebase ese campo lo alimenta el lector legacy (`LeerR1` toma la **primera** fila `Mes` de la fuente), no el agregado EXTEMP: julio-2 ASE5 = `13.084.233.937,44`, agosto-2 ASE5 = `-4.485,01`, frente a `441.849,00` / `1.249.810,00` del manual. El operando que produce los Δ del plan es el **agregado del dominio por firma** `WorkbookLeafInputsR1.ExtemporaneoEsperadoPorAse` (coincide al centavo con `CONSOLIDADO!D47:D51` del manual en julio-2 y agosto-2). `ConsolidadoAse.Extemp` queda **intacto** (fuera de alcance: el plan no corrige el EXTEMP).
- **Firma del falso verde (probada):** julio-2 → `D51 = 0` vs `441.849,00` (Δ `441.849,00`); agosto-2 → `D51 = 0` vs `1.249.810,00` (Δ `1.249.810,00`); ASE1-4 cierran exacto.

### A1-interior (acotado: ASE5, columna F)

- Mapa **DATO citado** del T0 §3.4: `Remuneracion.Core.Models.AnclasExtempR1` → ASE5 `{F682: F660+F629−L629; F702: F659+F628−L628}`. El error nombra **hoja R1 + celda ancla** (localiza dónde se perdió el importe).
- Tercera forma (ancla ausente con secciones presentes, o ancla con `<f>` inesperado) = **`ERR-PLANTILLA`** nombrando hoja+celda, nunca 0 silencioso ni fallback posicional. Una v3 exige re-congelar el mapa con su T0. Un workbook sin NINGUNA celda de esa geometría (p. ej. el manual legacy, fixture de lectura) no es tercera forma: no hay ancla que verificar.

### Regla celda-referenciada (doctrina de T0)

**Ninguna celda referenciada por otra hoja puede congelarse como "delta del manual" sin demostrar que el valor NO altera el consumidor.** La prueba de que `F682/F702` era un defecto no fue la diferencia de texto: fue que `CONSOLIDADO!D51`, `DetRetri!D36`, `DetValiRetri!D28` y `REMUNERACION_OCCIDENTE!D51` la consumen. Los 5 T0 que asertaban el literal están **reclasificados como DEUDA DECLARADA** (región `DEUDA-DECLARADA Plan 49`, asserts **byte-idénticos**: `GeometriaPlantillaEstandarizadaT0Tests`, `T0ResidualOccidenteAse5Tests`, `AdopcionPlantillaV2T0Plan48Tests`, `R1FuenteCanalConceptoT0Tests`, `DetRetriQ2Tests`). Condición de retiro: **v3 de la propietaria** con la fórmula restaurada + T0 nuevo.

### Estado de suite y rojo esperado

- Suite `p49-h2-suite.trx`: **1061 Total / 1012 Passed / 49 Failed / 0 Skipped** vs base `p48-h3-suite.trx` (1059/1021/38/0): **11 altas, 0 bajas**, todas `Plan 49 A1` sobre **julio-2** (detalle en `Remuneracion.IntegrationTests/TestResults/p49-h2-setdiff.txt`). No se re-congelaron ni se reverdecieron: el rojo es el éxito del plan hasta la v3.
- **Prohibido reverdecer** julio-2/agosto-2 tocando tests; el `12.462,26` del R2 ASE4 **no** se declara en este plan (T0 de dueño propio).

## Doctrina fuentes de recaudo opcionales + control de etiquetas (Plan 50 — vigente)

**Principio rector (`la fuente de recaudo ausente es un caso normal: 0 explícito con huella; la etiqueta desconocida NO`):** una fuente de recaudo (`ENEL`, `OCCIDENTE`, `NUEVO ESQUEMA`, `ENERBIT`) puede no venir en un período en R1, R2 o R4 de forma independiente. El reader materializa **0 por celda del mapa + huella auditable** y la corrida continúa; la **estructura rota** (invariantes de cierre T0e, fila `Componente`) sigue siendo fail-fast. Cero ramas por empresa/período: la ausencia se predica de la **geometría observada** (cero filas del label), no de una lista.

### A1 — opcionalidad (reader, W-A)

- **Ausencia total** (`filasTotal + filasSubs == 0` en R1, `filasEmpresa.Count == 0` en R2, sin fila `A=label B=Total` en R4) → todas las celdas del mapa por empresa = `0m`, un `Log.Warning` por (ASE, reporte, fuente) y entrada `{Reporte}:{label}` en `ConciliacionEmpresaInputs.FuentesAusentes` → propagada aditiva a `ResultadoProcesoPeriodo.FuentesAusentes` (`ASE n · R1:NUEVO ESQUEMA`).
- **Parcial con label presente** (el label trae filas pero falta el bloque que un operando exige, p. ej. `TotalMedio` = índice 1 con una sola fila `Total`) → **fail-fast vigente** (mensajes actuales de `ValorDe`/`EspecialesDe`). Sólo un T0 de negocio con bytes puede moverlo a 0 documentado. **Caso vivo declarado:** octubre-1 ASE3 R1 OCCIDENTE (una sola fila `Total`) — `Docs/T0-Plan50-FuentesRecaudoOpcionales.md` §2.2.
- **Writer intacto (W-A):** `OpenXmlPlantillaWriter.EscribirCeldasConciliacion` sigue exigiendo valor por celda del mapa; el 0 lo aporta el reader. El guard anti-fórmula (`ERR-PLANTILLA`) **no se relaja**.
- **Precedente:** `AjustesSfT` con `EsOpcional` y `Bloque5Total → 0` (el mapa por empresa **no** gana flags).

### A2 — Q2: recorte MANTENIDO por residuo de plantilla (`R-MAPA-V2-R2-ASE4`)

El recorte de la conciliación por empresa en Q2 (`ProcesadorPeriodo`, HU-20-T0b/G2-D1) **sigue vigente**: al intentar levantarlo, la escritura de la celda destino `Rem. Anticipos R2!E285` (mapa `(1,4)`) aborta con **`ERR-PLANTILLA`** porque en la **v2** `E285 = E192+E203-K192` es fórmula (en la v1 era literal 0), igual que `K285`. Reactivarlo exige **re-congelar ese mapa por T0** (clase Plan 48), nunca editar el mapa ni relajar el guard. Evidencia: `Docs/T0-Plan50-FuentesRecaudoOpcionales.md` §4.3. **A1 es period-free**: el reader lee igual en Q1 y Q2; lo que se omite en Q2 es la llamada del procesador.

### A3 — control anti-falso-0 = CONTROL DE ETIQUETAS (no aritmética)

El T0 midió que `Σ fuentes presentes vs total del reporte` **no es viable**: en R1 el total del cuadro `Componente` no es la suma de las filas-etiqueta (Δ 2.015,07 / 243.714,65 / 558,05 / 30.180,94 en octubre-1; tampoco 0 en julio-2/agosto-2) y en R2 el total incluye la etiqueta real `EAB` (sin destino en el mapa) → Σ-reconocidas ≠ total. En R4 la igualdad es una tautología (la fila `Total` es la suma de las filas-etiqueta). Control implementado:

- **Evidencia (reader, sin juicio):** `ConciliacionEmpresaInputs.EtiquetasNoReconocidas` = etiquetas observadas en las filas de cierre (`R1:C` con `D=Total`, `R2:B` con `C=Total`, `R4:A` con `B=Total`) que **no** pertenecen al vocabulario de fuentes (`EmpresaFacturacion.Catalogo` + rótulo vigente por ASE + sinonimia `RECIPROCIDAD ↔ NUEVO ESQUEMA`), formato `{Reporte}:{label}`.
- **Juicio (Core):** `ValidadorBasico.ValidarFuentesRecaudoNoReconocidas` → **fail-fast** sólo si una etiqueta no reconocida coexiste con una **fuente ausente en el MISMO reporte** del mismo ASE (posible fuente presente no reconocida). Sin ausencia en ese reporte la etiqueta es **sólo evidencia** (hoy: `EAB` en R2 de ASE2/3/4 — deuda del Plan 49).
- **Sin umbral numérico** (no hay resta): el control es de etiquetas. Es el mínimo alternativo del §2.5 del plan, con su T0 propio.

### Residuo declarado del gate Σ (`R-SIGMA-Q2`)

`Σ visibles-empresa == visible de bloque` (HU-08 §2.5: `ValidadorBasico.ValidarSigmaEmpresasPorAse` y el gemelo `WorkbookLeafCoherence.ValidarSigmaEmpresas` del writer) se definió sobre los mapas Q1 y **no descompone** el visible de bloque en Q2: R4 sólo declara empresas con celda destino (`ENEL` sin celda en ASE1/3/5), R2 omite etiquetas reales sin destino (`EAB`, `ENERBIT`) y el visible R1-Q2 se deriva por firma. Queda **acotado a Q1** (comportamiento Q2 idéntico al anterior). Corregirlo exige contrato nuevo + cobertura del mapa (T0 propio).

### Regresión y evidencia

- T0 read-only: `Remuneracion.IntegrationTests/FuentesRecaudoOpcionalesT0Plan50Tests.cs` (matriz período×ASE×reporte, parcial, aritmética A3 y etiquetas) + `Docs/T0-Plan50-FuentesRecaudoOpcionales.md`.
- H1: `FuentesAusentesPlan50Tests` (ausencia→0+huella, parcial→fail-fast, invariante→fail-fast, period-free) · `ControlAntiFalsoCeroPlan50Tests` (A3).
- **Línea base viva del árbol** (post-commit `d55e621`, que reemplazó la v2): `TestResults/p50-h0-baseline.trx` = 1061/994/67/0 — 18 rojos más que `p49-h2-suite.trx` (1061/1012/49/0) por el cambio de plantilla, **antes** del Plan 50. El set-diff de H2 se hace contra esa base.

## Doctrina reporte por banco: columnas por HEADER + fuente ausente omitida (Plan 51 — vigente)

**Principio rector (`la columna de la fuente del banco se reconoce por su ETIQUETA, no por su índice; la fuente que no viene se omite con huella`):** el Resumen del `ReportePagosxBanco` trae un header de empresas cuya **posición física cambia por ASE y por período** (ASE1/3: `ENEL[1] OCCIDENTE[2] Total[3]`; ASE2/4: `ENEL[1] NUEVO ESQUEMA[2] OCCIDENTE[3] Total[4]`; ASE5: `ENEL[1] ENERBIT[2] OCCIDENTE[3] Total[4]`; en octubre-1 ASE2 y ASE4 vienen SIN `NUEVO ESQUEMA` y `OCCIDENTE` baja a `[2]`). El mapa **posicional** del T0-0.3 abortaba con `la columna fuente 2 … debería ser 'NUEVO ESQUEMA' pero dice 'OCCIDENTE'`; quedó **retirado** (`ColumnasFuentePorAse` ya no existe).

### A1 — header-scan (reader `LeerReporteBanco`)

- Se recorre la fila de headers (columna `A` = etiqueta de concepto, nunca fuente). Cada celda **no vacía** y cuyo normalizado **≠ `total`** se resuelve contra el **vocabulario de fuentes del negocio**: los 5 `LabelTemplate` del catálogo + el rótulo vigente por ASE (`ObtenerLabelFuente`: empresa 1 en ASE2/4 → `NUEVO ESQUEMA`) + la sinonimia citada `RECIPROCIDAD ↔ NUEVO ESQUEMA` (Plan 34), vía `SinonimosEmpresaR1.SonMismaEmpresa`. La clave resultante es la de la **plantilla** (`EditablesPorAse[ase].Keys`), no un catálogo nuevo.
- **Lo esperado por ASE lo declara la plantilla**: `EditablesPorAse[ase].Keys`. `Total` se excluye **antes** del vocabulario; las celdas vacías se ignoran.
- **Ausente = esperada no observada** → la fuente **no se emite** (no se leen sus conceptos, el writer itera presentes y sus celdas quedan en el **0 de la plantilla**) + huella (A2). No hay `TotalFuente = 0` que confundir con un total real y la fuente ausente **no entra** al gate Σ por empresa.
- **Sin ramas por período/empresa/fuente** (D-H): la ausencia se predica de la geometría observada (headers del resumen). El orden de columnas puede invertirse y el mapeo sigue correcto.

### A2 — huella `BANCO:` (canal durable existente)

`ReporteBancoAseInputs.FuentesAusentesBanco` = `BANCO:{label}` (p. ej. `BANCO:NUEVO ESQUEMA`) → `ResultadoProcesoPeriodo.FuentesAusentes` como `ASE {n} · BANCO:{label}` + **1 `Warning` por (ASE, fuente)** en lenguaje administrativo. Grano **ASE** (el banco no se replica ×5 como la conciliación por empresa). El prefijo `BANCO:` no contamina la partición por reporte R1/R2/R4 del Plan 50 (`ValidarFuentesRecaudoNoReconocidas` lee solo `leaf.Conciliacion`).

### A3-banco — etiquetas, sin umbral (juicio en Core)

`ReporteBancoAseInputs.EtiquetasBancoNoReconocidas` = `BANCO:{label}` (evidencia del reader) + `ValidadorBasico.ValidarEtiquetasBancoNoReconocidas` (juicio, junto al control del Plan 50, heredado por `ClasificarDivergenciasDeclaradas`):

- **desconocida + ausente en el MISMO ASE** → fail-fast `ASE n · BANCO · control de fuentes: …` (firma de "fuente presente no reconocida", cambio de etiqueta), nunca 0 silencioso;
- **desconocida sin ausencia** → sólo evidencia (no convierte en rojo lo sano);
- **cero fuentes reconocidas** → fail-fast (estructura rota);
- **fuente reconocida sin celdas destino** en `EditablesPorAse[ase]` → fail-fast (desync mapa↔plantilla, bug real);
- **sin umbral numérico**: el gate Σ por empresa presente sigue intacto y la ausente no entra a él.

### Superficie intacta

`OpenXmlPlantillaWriter.EscribirCeldasBanco` **no cambia** (itera `bloque.Empresas`, consulta `editables.TryGetValue` solo para las presentes y escribe C59 siempre); plantillas v1/v2 intactas; conceptos 1/2/3/7 y fila `Total` se leen igual (prefijos + `total` intactos); cero fórmulas tocadas.

### Evidencia

- T0 read-only: `Remuneracion.IntegrationTests/ReporteBancoHeadersT0Plan51Tests.cs` (105 casos: matriz 5 ASE × 4 períodos, `Total`, conceptos, Σ, paridad, unknowns) + `Docs/T0-Plan51-Headers-Banco.md`.
- H1: `ReporteBancoFuentesOpcionalesPlan51Tests` (10) · H2: `ControlBancoNoReconocidoPlan51Tests` (4).
- Suite `p51-h2-suite.trx` = **1260 / 1193 / 67 / 0** vs base `p50-h2-suite.trx` (1141/1074/67/0): **119 altas (solo tests Plan 51), 0 bajas, 0 rojos nuevos**. `TestResults/p51-h2-setdiff.txt`.
- **Desviación declarada del T0:** el plan esperaba que sólo ASE2 llegara con 2 fuentes en octubre-1; la medición mostró que **ASE4 también** (misma familia de ausencia legítima, absorbida por la regla genérica).

## Doctrina contrato de SECCIÓN por operando/celda R1 (Plan 53 — vigente, fin del ordinal)

**Principio rector (`la celda destino declara QUÉ SECCIÓN la alimenta; la ocurrencia y la magnitud quedan prohibidas`):** el mapa `WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa` deja de resolverse por ocurrencia sobre listas ordenadas por `|F|` (`Main=[0]`, `Medio=[1]`, `Menor=[count-1]` — el defecto que el Plan 52 §13 cerró con STOP) y pasa a declarar **tres campos por celda**: `(Celda, FuenteR1, SeccionEsperadaR1)`.

- **Roles relativos** (`Remuneracion.Core.Models.SeccionEsperadaR1`): `UltimaOportuna` (última sección del bloque de la empresa), `Primera` (la primera **sólo si hay ≥2**: con una sola sección no existe "el otro" bloque), `Intermedia` (la de en medio **sólo si hay ≥3**, D4), `UnicaRecaudo` (la única sección de recaudo; misma resolución que `UltimaOportuna`, declarada donde el T0 probó la geometría de sección única y **prohíbe** la lectura "única primera ⇒ Menor" de la regla D3-del-52, derogada).
- **Resolvedor puro** `Remuneracion.Core.Models.ResolvedorOperandosEmpresaR1` (sin I/O, sin período, sin empresa — D-H): recibe las filas-empresa **ya tipadas** por el predicado vigente `ContratoSeccionPeriodFreeR1.SeccionPorCierre` (Plan 46) y devuelve `fila | ausente | empate`. Sección declarada **ausente → `0m`** (la fuente no la trae: es normal, reglas 6 y 7); **empate** (2 filas-empresa en la misma sección) → `CalculoInvalidoException` accionable con ASE + empresa + sección + filas (no hay regla de desempate aprobada).
- **Punto único de lectura**: `ExcelDataReaderWorkbookLeafInputReader.ConstruirSecuenciaEspejoR1` (ventana `E='Vlr Servicio'` → `A='Total'` con `B` vacío + índice crudo por fila) lo comparten `LeerEspejoR1` y la conciliación por empresa (`ExtraerCeldasR1` en el Plan 53, **retirado en el Plan 55** y reemplazado por `ResolvedorConciliacionEmpresa`, que recibe la MISMA secuencia); prohibido reimplementar el scan de fronteras. Columnas de lectura intactas (F=5, Especiales=11); `Block5Total` (canal `D=5`) conserva semántica sin sección.
- **Contrato derivado, no inventado**: 56 entradas congeladas celda a celda en `Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md` §3 (sección de la fuente + `<f>` consumidora donde existe + oráculo del manual del administrativo, 4/4 al centavo). El T0 midió **180/180 filas-empresa tipadas**, 0 empates, 0 intercalados y formas de ≤3 secciones.
- **Evidencia y gates**: T0 read-only `SeccionPorOperandoR1T0Tests` (11) · unit del resolvedor `ResolvedorOperandosEmpresaR1Tests` (10) · contrato con insumos reales `ContratoSeccionPorOperandoR1Tests` (8, incluye no-duplicación oct-1 ASE3 OCCIDENTE y no-pérdida oct-1 ASE5 ENERBIT) · reescritura de `FuentesAusentesPlan50Tests.S2_1_ParcialConLabelPresente` a aserciones de valor (el residuo `R-PARCIAL-OCT-ASE3-R1-OCC` queda cerrado).
- **Set-diff `p53-h2-setdiff.txt`**: **29 altas (solo tests Plan 53), 0 bajas, 0 cambios de resultado** contra `p51-h2-suite.trx`; ningún rojo nuevo. **Desviaciones declaradas:** (a) la pata "geometría del bloque oficial" de D2 no existe para 34/56 celdas (slots de valor de la v2 + direcciones fuera del bloque detalle; condición preexistente del frente E1-2) — la tipabilidad se sostiene en la sección de la fuente + el `<f>` donde existe + el manual; (b) la lista cerrada suma 2 Δ (`F474`/`F477` de octubre-1 ASE5) que §9 no enumeraba, normados por D5 (misma regla que §9 #1/#2/#5/#8/#12).

## Doctrina contrato ÚNICO de lectura R1 (Plan 54 — vigente, fin de la doble resolución)

**Principio rector (`la quincena NO selecciona algoritmo`):** la lectura R1 se resuelve con **un solo camino** para Q1 y Q2 — roles por **firma de etiquetas (A/B/C/D/E)** sobre la secuencia tipada (`BloqueEspejoAseInputs.Filas`), un solo método `ExcelDataReaderWorkbookLeafInputReader.MapearR1`. `Periodo.NumeroQuincena` gobierna **solo fuentes y declaración de visibles** (detalle R2/R4 solo-Q2, `AjustesSfT = 0` en Q1, columnas D/E-vs-F/G del RESUMEN MES, DetRetri Q2-escribe/Q1-valida, naming, preflight Q2-only y el recorte §7), **nunca** la resolución del leaf. Este apartado **supersede** el dispatch por quincena y la cardinalidad `Mes0/1/2` core del Plan 25; el resto de su doctrina (firma, D-A, D-C, D-F) sigue vigente.

- **Mapa R1 productivo ÚNICO** = `WorkbookLeafCellMapQ2.R1Q2EditablesPorAse` (geometría v2), para ambas quincenas. `WorkbookLeafCellMapPorAse.EditableLeafCellsPorAse` queda **LEGACY (geometría v1)**: solo lo consumen el API histórico `EscribirDetalleR1` (falla honestamente con `ERR-PAYLOAD`, fuera de paridad) y fixtures/tests de caracterización. Prohibido re-apuntar producción a la v1. R2/R4 también colapsaron a **una tabla destino única** (`R2Q2EditablesPorAse` / `R4Q2EditablesPorAse`).
- **Regla de ausencia (slot declarado que la fuente no llena → CERO del template, nunca aborto):** `Mes0`/`Mes1` (y `Lmes0`/`Lmes1`) son **core** con fail-fast que nombra ASE + reporte + celda (mensaje intacto); `Mes2`/`Lmes2` y `Aplic0`/`Aplic1`/`LAplic0` son **opcionales → `0m`**. Evidencia: julio-1 ASE1 trae 2 filas `Mes/Total` y su R10 oficial se reproduce al peso con `F70 = 0` (`16.758.167.585`); la ausencia **no** se generaliza a cualquier rol (la clase completa "slot declarado sin fila fuente" con control anti-falso-0 es del Plan 55). `F30`/`F10`/`L10` son **constantes `0m`** (miembros conservados por contrato, sin lector productivo).
- **EXTEMP v2** = `Σ F(Aplic) − Especiales(Aplic[0])`, **sin término de subsidio**: el agregado `F30+F10` (subsidio[0] + Aplic[0]) era **semántica v1** y quedó erradicado con la rama Q1. `totOpt` = `Σ F(Mes) − Especiales(todas menos la última)`, intacto (Plan 25 D-F).
- **Candado anti-recaída** `Remuneracion.IntegrationTests/LecturaR1SinRamaQuincenaTests.cs`, dos capas: **estática** (el lector leaf no puede contener `NumeroQuincena` fuera de la allow-list cerrada de usos de D7, y el test enrojece ante un uso injertado) + **conductual** (la MISMA fuente con `Periodo` Q1 y Q2 produce `WorkbookLeafInputsR1` idéntico miembro a miembro: `F25/F41/L25/F30/F10/L10`, ambos agregados y `CeldasPorAse`).
- **Divergencias declaradas:** `ASE5-F674-MES2-AUSENTE` en `DivergenciasDeclaradasCatalogo` (estructural, **Δ = 0**: la fuente trae 2 filas `Mes/Total`, el slot `F674` queda en 0 y `F680` cierra contra el R10 ⇒ documenta y **no** matchea, porque no hay discrepancia material). Residuos vivos: **`R-RECORTE-Q2`** (`ProcesadorPeriodo.cs:198-217`, conciliación por empresa omitida en Q2 por `R-MAPA-V2-R2-ASE4`) y **`R-ANCLA-EXTEMP`** (`AnclasExtempR1` congela `F682`/`F702` como literales, mientras los bytes de la v2 traen `F682 = <f>F660+F629-L629</f>`; produce `ERR-PLANTILLA` en las corridas y es rojo de base).
- **Evidencia y gates:** T0 read-only `ContratoUnicoLecturaR1T0Tests` (58) + `Docs/T0-Plan54-Contrato-Unico-Lectura-R1.md`; H1 `ContratoUnicoLecturaR1H1Tests` (16); H2 `LecturaR1SinRamaQuincenaTests` (6) + `ContratoUnicoLecturaR1H2E2ETests` (5). Suite `p54-H2-suite-final.trx` = **1374 / 1307 / 67 / 0** vs base `p53-h2-suite.trx` (1289/1222/67/0): **+85 altas verdes, 0 bajas, los 67 rojos son el MISMO conjunto (0 nuevos, 0 desaparecidos)**; build 0 warnings; SHAs v1/v2 intactos; cero `<f>`/`InsertRow`/`DeleteRow`.

## Doctrina conciliación por empresa v2 — resolvedor por etiquetas (Plan 55 — vigente; supersede la tabla congelada)

**Principio rector (`la plantilla es la autoridad`):** el camino por empresa (R1/R2/R4) resuelve cada destino por **cruce de etiquetas** contra el cuadro del template vigente: `(cuadro/sección, fuente = col C, canal = col D, concepto = col E) → celda`, con el **componente resuelto por el ENCABEZADO de columna** que declara el propio cuadro. Maquinaria reusada (no reinventada): `SinonimosEmpresaR1` (catálogo abierto), `ContratoSeccionPeriodFreeR1.SeccionPorCierre`/`CanalEfectivo`/`Alinear` y las columnas por encabezado del camino de bloque. Unidad única: `Remuneracion.Infrastructure/Excel/ResolvedorConciliacionEmpresa.cs` (`Resolver`), alimentada por `ExcelDataReaderWorkbookLeafInputReader.LeerConciliacionEmpresas`.

- **Prohibido en el camino de resolución:** listas por ASE, literales de empresa, tablas congeladas por `(empresa, ASE)` y ramas por quincena. **Candado estático** `Remuneracion.IntegrationTests/ConciliacionEmpresaCandadoAntiQuemadoTests.cs`: allow-list cerrada de archivos/regiones + prueba de injerto (una rama por ASE, un literal de empresa o el mapa congelado inyectados a propósito enrojecen el test). El vocabulario de empresas sale del **catálogo + sinonimia**; una etiqueta que el cuadro declara y el catálogo no cubre (p. ej. `EAB`) entra como empresa de fuente (Id sintético).
- **Ausencia** (el cuadro declara `(fuente, canal, concepto)` y la fuente no lo trae) → **cero del template**, sin error, con huella `{Reporte}:{etiqueta}` **solo donde el cuadro declara el rótulo** (si no lo declara, no aplica: no es ausencia). Los invariantes de estructura del reporte (`Componente/Total`, etc.) siguen siendo fail-fast.
- **Exceso** (la fuente trae algo que ningún cuadro declara) → **registro durable** en `ConciliacionEmpresaInputs.DestinosSinCelda` + `Log.Warning` + divergencia declarada (Plan 44/Plan 42): nunca descarte silencioso, **nunca celda inventada**, nunca aborto por defecto. Celda destino con `<f>` → no se escribe y se registra en `DestinosFormula` (guard anti-fórmula).
- **Agregado por empresa ANTI-TAUTOLÓGICO:** se deriva de la **FUENTE**, agrupado por etiqueta, sobre el MISMO conjunto de secciones/roles que alimenta el visible de bloque — R1: `Σ Total(secciones Oportuno) − Σ Especiales(todas menos la última)`; R2: `Componente + Subs/Cont − Especiales(Componente)`; R4: `Total`. **Prohibido** definirlo como "la suma de lo que escribimos": el gate debe seguir siendo una verificación independiente del camino de escritura (test `H2_1_CA8_...`).
- **`WorkbookLeafCellMapPorEmpresa` = LEGACY** (mismo patrón que `WorkbookLeafCellMapPorAse` del Plan 54): fuera del camino productivo, solo lo consumen fixtures históricos; el writer pega únicamente las celdas que el reader resolvió.
- **Residuo vivo del bloque de cierre:** `R-SLOT-Q1-ENERBIT` — la corrida 5-ASE de `2026101` cierra el gate Sigma y el DetRetri contra el R10, y **después** aborta en `ERR-SIN-SLOT` de `EscritorValoresSlotR1` (ASE5, fuente R1, rótulo `ENERBIT`, canal `Total`, `1.ª sección`, ocurrencia #1): deuda pre-existente de la familia E1-6 (Planes 41/46), no agravada por el Plan 55. Gate: plan propio que revise el contrato de sección Q1 ↔ slots del template, **sin** remap/insert/posición.
- **Evidencia:** `Docs/T0-Plan55-Conciliacion-Por-Empresa.md` (H0, 81/81) · `ConciliacionEmpresaSigmaPlan55Tests` (gate por astilla + D1-D5 + CA8 + `DetRetri` = R10) · `ConciliacionEmpresaCandadoAntiQuemadoTests` (T4). Suite `p55-H3-suite.trx` = **1484 / 1417 / 67 / 0**, set-diff de rojos **vacío** vs `p54-H3-suite.trx` y `p53-h2-suite.trx`; build 0 warnings; SHAs v1/v2 intactos; cero `<f>`/`InsertRow`/`DeleteRow`.
