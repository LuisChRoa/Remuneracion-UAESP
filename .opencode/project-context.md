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
- `{periodo}/Conciliaciones/` — 5 archivos `Conjunta {prefijo}*.xlsx` / `Directa*.xlsx` (hoja única `RESUMEN MES`). Fuente de las hojas `Recaudo *` (G1: `BuscarConciliacion` resuelve SOLO esta ruta, sin fallback).
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

**Principio rector (`validar TODO antes de procesar`):** ante un período con insumos incompletos, el flujo aborta en segundos —antes de abrir cualquier workbook— con UN solo error `ERR-FUENTE-NO-ENCONTRADA` (salida 2) que enumera **de una vez TODOS los faltantes**, en vez de procesar 5 ASE para descubrir al final un faltante por corrida. El preflight es una puerta de **existencia en filesystem**: NO valida contenido (headers/estructura/coherencia numérica siguen fallando en sus readers como hoy).

**Superficie verificada = superficie que el runtime resuelve (ni un archivo más ni uno menos):** por ASE, R1 + R2 + R4 + banco + balance SIEMPRE; `SaldosaFavorAplicadosPorNotas` y `RetribuciónNegativa` SOLO si `Periodo.NumeroQuincena == 2` (gobierno por dominio, nunca por detección de contenido). De período: las **5 conciliaciones SIEMPRE (Q1 y Q2)** — la lectura `Recaudo *` está fuera de toda rama por quincena (§V3) — + R10 SIEMPRE (oráculo obligatorio G3 en ambas quincenas). `RecaudosReversados` queda **FUERA**: el runtime no lo consume (sin finder ni lectura, §V9); exigirlo sería inventar un requisito.

**Finders intactos; el preflight valida, no resuelve (D-C):** `ValidadorInsumosPeriodo` REUTILIZA los finders de `ILocalizadorArchivosAse` (mismo prefijo case-insensitive, misma normalización de diacríticos, mismo `TopDirectoryOnly`); no duplica lógica de match ni reemplaza la resolución runtime. Los `?? throw` intermedios de los procesadores **permanecen** como defensa en profundidad. Si UAESP agrega o retira un insumo que el runtime resuelve, se agrega/retira en la lista del validador con su finder (checklist de paridad preflight↔runtime).

**Mensaje en lenguaje administrativo (D-E):** encabezado `Faltan insumos para el período {AAAAMMQ} (quincena {1|2}). No se procesó ningún ASE.` + lista numerada; cada ítem dice QUÉ falta (nombre que el usuario reconoce: ASE, reporte, carpeta), DÓNDE debe ir (ruta y nombre exacto o inicio de nombre esperado) y QUÉ hacer (solicitar/generar, colocar, reejecutar). **Jerga de código PROHIBIDA** en el texto de los ítems (`prefijo`, `matcher`, `finder`, `TopDirectoryOnly`, nombres de clase o código interno). Si falta la carpeta `Conciliaciones/` completa, UN solo ítem la nombra + enumera los 5 archivos esperados (no 6 líneas redundantes); igual para una carpeta ASE ausente. Orden estable: ASE 1..5 (carpeta, luego reportes en el orden R1, R2, R4, banco, balance, [Q2] saldos-notas, retribución), luego `Conciliaciones/`, luego R10.

**Integración:** `ProcesadorPeriodo.Ejecutar` corre el preflight tras resolver las 5 carpetas ASE y ANTES del loop por ASE (`progreso?.Report` + `Log.Warning`; doctrina HU-14); con faltantes NO se lee, no se calcula y no se escribe. `ProcesadorRemuneracion.Ejecutar` (single-ASE) solo verifica existencia (`File.Exists`) de R1/R2/R4 — guardrail mínimo; NUNCA exige banco/balance/conciliaciones/R10/saldos. Sin faltantes el flujo es bit-idéntico al actual (cero-geometría: ningún valor ni fórmula cambia; el diff es el servicio/modelo Core + 2 puntos de llamada).

**Follow-up explícito (fuera del Plan 26):** validación de CONTENIDO anticipada (headers/estructura antes de procesar) requiere su propio T0: abre workbooks, bloquea archivos y duplica readers (descartado por diseño, no por olvido).

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
