# AGENTS.md — Remuneración Quincenal UAESP

## Project Context

Automates bi-weekly waste management remuneration for Bogotá (UAESP), distributing payments among 5 operators (ASE) per Colombian regulation Resolución 27 de 2018. Source of truth for the business spec: `README.md` and `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md`.

**Current state:** 3-layer solution is *scaffolded* (interfaces, models, constants, exceptions, FileSystem locator) but the Excel I/O implementations are **STUBS that throw `NotImplementedException`**. The next real work is implementing those stubs against real source files. **There is no test project yet.**

## Tech Stack

| Layer | TargetFramework | Packages |
|-------|-----------------|----------|
| `Remuneracion.WinForms` (UI, `net10.0-windows`) | WinForms, `-windows` only here | Serilog 4.4.0 + Serilog.Sinks.File 7.0.0 |
| `Remuneracion.Infrastructure` (`net10.0`) | Excel | ExcelDataReader 3.9.0 (read), DocumentFormat.OpenXml 3.5.1 (write) |
| `Remuneracion.Core` (`net10.0`) | none (pure domain) | — |

All packages are already in the `.csproj` files — **do not re-add them** with `dotnet add` unless changing versions. Suggested test stack (README plan): xUnit + FluentAssertions — none present yet.

## Architecture & Wired Entrypoints

```
Remuneracion.WinForms (UI) ──► Remuneracion.Infrastructure (Excel I/O)
        │                            │
        └────────────────────────────┴──► Remuneracion.Core (domain, no deps)
```

Defined contracts in `Core/Interfaces/` (implementations live in `Infrastructure/`):
- `IRecaudoReader` — `LeerR1/LeerR2/LeerR4` → impl `ExcelDataReaderRecaudoReader` (STUB)
- `IPlantillaWriter` — `EscribirConsolidado/EscribirDetalleR1/R2/R4` → impl `OpenXmlPlantillaWriter` (STUB)
- `ICalculoRemuneracion`, `IValidador` — **no implementations exist yet**; create them in Core when implementing
- `Infrastructure/FileSystem/ArchivoFuenteLocator` — already implemented (locates ASE folders by `N-` prefix and files by `Recaudopor...`/source prefix)

Domain models in `Core/Models/`: `Ase`, `Periodo`, `RecaudoComponenteR1`, `SaldosFavorR2`, `ReversionR4`, `ConsolidadoAse`, `ResultadoRemuneracion`. Note `ConsolidadoAse.TotalAse`, `Periodo.CodigoCompleto`/`NombreArchivo`, and `SaldosFavorR2.TotalOportuno` are **computed** getters — keep ASTs in sync with the README formulas.

## Development Commands

```bash
# Build (solution is inside Remuneracion.WinForms/)
dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx

# Run the WinForms app (WinExe — no console output)
dotnet run --project Remuneracion.WinForms/Remuneracion.WinForms.csproj

# Logs
# Serilog writes to ./remuneracion_log.txt (gitignored)
```

There is no test project / no CI / no lint/format config. To add tests, create a new `net10.0` xUnit project referencing `Remuneracion.Core` (and `Remuneracion.Infrastructure` for I/O tests) and add it to the `.slnx`.

## Critical Business Rules

These affect code directly and are easy to get wrong:

- **NEVER overwrite formulas** in the Excel template — paste only values
- **ASE4 (Bogotá Limpia)** has no "Especiales" column — check column headers dynamically (`SaldosFavorR2.TieneColumnaEspeciales` / `ServEspK`; if absent use 0)
- **SALDOS POR NOTA** and **RETRIBUCION NEGATIVA** only apply to the 2nd fortnight (`Periodo.NumeroQuincena == 2`); `AjustesSfT` is 0 in quincena 1
- **Tolerance:** differences ≤ ±0.5 between calculated and source values are acceptable (rounding)
- **Dynamic row positions:** source file rows aren't fixed — search by header values, not row numbers
- **R4 reversal values** are stored negative in column 4
- **Round to integer** before writing values to CONSOLIDADO column D (`DetRetri`)

## Cell Reference Map (CONSOLIDADO)

| Cells | Calculation | Source File |
|-------|-------------|-------------|
| D9:D13 | TOT_OPT | Recaudoporcomponente (R1) — row where col1="Componente", col2="Total", col F |
| D28:D32 | R2 Total Oportuno | RerpoteDetalleSaldosaFavor (R2) — Grand Total (E) − SERV_ESP_K |
| D47:D51 | EXTEMP | Recaudoporcomponente (R1) — Extemporáneo section, col F |
| D66:D70 | Reversión R4 | ReversiónPorComponente (R4) — last row, col 4 (negative) |
| D85:D89 | AJUSTES-SF-T | SALDOS POR NOTA + RETRIBUCION NEGATIVA (2nd fortnight only) |
| D104:D108 | Total por ASE | Sum of above |
| D109 | Gran Total | SUM(D104:D108) |

## Real Data — `Docs/Insumos/`

Contains real source/template files that the stubs require (currently **untracked in git** — do not assume they're committed):
- `Remuneracion 202607-1 Total.xlsx` / `Remuneracion 202607-2 Total.xlsx` — completed output templates for period 202607 (reference for writing)
- `A1 _ R4-BalanceSubsidioyContribuciones_..._20260724...xlsx` — source R4/Balance file

Use these as fixtures when implementing/validating the readers and writers. Follow the source-file naming pattern: `Recaudoporcomponente_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{timestamp}.xlsx` (and the `RerpoteDetalleSaldosaFavor_*` / `ReversiónPorComponente_*` variants).

## Directory Layout (current)

```
Automatización/
├── README.md                       # Full business spec — READ THIS FIRST
├── Docs/
│   ├── Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md  # Architecture & phases
│   ├── Detalle de plantilla.docx        # Sheet-by-sheet instructions
│   ├── Proceso de Recaudo.docx          # Payment receipt flow
│   ├── Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx    # Executable prompt
│   └── Insumos/                         # Real source/template xlsx (NOT committed)
├── Remuneracion.Core/               # net10.0, pure domain (Models, Interfaces, Constants, Exceptions)
├── Remuneracion.Infrastructure/     # net10.0, Excel I/O (ExcelDataReader, OpenXML) + FileSystem
└── Remuneracion.WinForms/           # net10.0-windows WinForms UI; contains the .slnx
```

## Phase Roadmap

1. **Fase 1 (Current):** Single ASE prototype — R1, R2, R4 reading + CONSOLIDADO calculation. Excel I/O stubs must be replaced with real implementations here.
2. **Fase 2:** Scale to all 5 ASE, add conciliation sheets, bank reports
3. **Fase 3:** Error handling, CLI mode, full logging, documentation

---

## Agent Architecture (Local — `.opencode/`)

> Basado en `Create-agent-architecture-plan_v3.1.md`. Agente principal **local** (auto-contenido en el repo); skills SIEMPRE globales (`~/.config/opencode/skills/`).

### Agente Principal
- **`remuneracion-architect`** — `.opencode/agents/remuneracion-architect.md`
  - Orquestador + mentor técnico + guardián de calidad. Delega, no implementa en primera persona cuando hay plan aprobado.
  - Regla de trato: dirigirse al usuario solo como **Inge** / **Ingeniero**.
  - Comunicación en español. Commits solo con `#commit`/`#push` explícito.

### Subagentes (según Marco de Decisión — calidad senior + Desktop .NET)
| Subagente | Archivo | Rol |
|-----------|---------|-----|
| `sdd-planner` | `.opencode/agents/sdd-planner.md` | Flujo SDD PROPOSE→DESIGN→SPEC→TASKS; plan en `plans/` |
| `implementer` | `.opencode/agents/implementer.md` | Implementación contract-first desde el plan aprobado |
| `code-reviewer` | `.opencode/agents/code-reviewer.md` | Spec compliance + calidad, 0 warnings |
| `debug-agent` | `.opencode/agents/debug-agent.md` | Causa raíz sistemática |
| `ui-ux-specialist` | `.opencode/agents/ui-ux-specialist.md` | Diseño premium WinForms + AI Slop detection |

> `dba-reviewer` NO creado: no hay base de datos (persistencia = archivos Excel). Aplica Regla #17 si apareciera un MCP de BD en el futuro.

### Datos de proyecto
- `.opencode/project-context.md` — stack, comandos, reglas de negocio, cell map, convenciones (inyectado junto con skills al delegar).
- `.opencode/opencode.jsonc` — permisos del agente principal (subtareas permitidas).
- `plans/` — planes SDD (fuente única de verdad para implementación).
- `requirements/` — requerimientos formales cuando la complejidad lo exige.
