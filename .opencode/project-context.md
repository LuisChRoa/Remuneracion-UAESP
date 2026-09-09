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
Archivos INSUMOS (NO commitados): templates completados `Remuneracion 202607-1/2 Total.xlsx`, fuente R4. Naming: `Recaudoporcomponente_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{ts}.xlsx` (+ variantes `RerpoteDetalleSaldosaFavor_*` / `ReversiónPorComponente_*`).

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
