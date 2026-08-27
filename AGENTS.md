# AGENTS.md — Remuneración Quincenal UAESP

## Project Context

Automates bi-weekly waste management remuneration for Bogotá (UAESP), distributing payments among 5 operators (ASE) per Colombian regulation Resolución 27 de 2018.

**Current state:** Early stage — skeleton WinForms app with no business logic. The `README.md` and `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` contain the full business specification.

## Tech Stack

| Layer | Technology |
|-------|------------|
| Framework | .NET 10 (net10.0-windows) |
| UI | Windows Forms |
| Excel Read | ExcelDataReader (MIT) — not yet added |
| Excel Write | OpenXML SDK (MIT) — not yet added |
| Testing | xUnit + FluentAssertions (planned) |

**Note:** NuGet packages for Excel handling are specified in the proposal but not yet in `.csproj`. Add them when implementing.

## Architecture (3-Layer)

```
RemuneracionQuincenal (WinForms UI)
  └── Remuneracion.Core (business logic, models, calculators)
  └── Remuneracion.Infrastructure (Excel read/write)
```

The current repo has only the UI project. Create Core and Infrastructure projects when implementing.

## Critical Business Rules

These affect code directly and are easy to get wrong:

- **NEVER overwrite formulas** in the Excel template — paste only values
- **ASE4 (Bogotá Limpia)** has no "Especiales" column — check column headers dynamically
- **SALDOS POR NOTA** and **RETRIBUCION NEGATIVA** only apply to the 2nd fortnight (AAAAMM2)
- **Tolerance:** differences ≤ ±0.5 between calculated and source values are acceptable (rounding)
- **File naming:** source files follow pattern `Recaudoporcomponente_to_date{INI}ddMMyyyy_to_date{FIN}ddMMyyyy___{timestamp}.xlsx`

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

## Directory Layout

```
Automatización/
├── README.md              # Full business spec — READ THIS FIRST
├── AGENTS.md              # This file
├── Docs/
│   ├── Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md  # Architecture & phases
│   ├── Detalle de plantilla.docx          # Sheet-by-sheet instructions
│   ├── Proceso de Recaudo.docx            # Payment receipt flow
│   └── Prompt_Maestro_Proyecto_Remuneracion_UAESP _ Vo .docx   # Executable prompt
└── RemuneracionQuincenal/
    ├── RemuneracionQuincenal.slnx         # Solution file
    ├── RemuneracionQuincenal.csproj       # .NET 10 WinForms
    ├── Program.cs                         # Entry point
    ├── Form1.cs                           # Main form (empty)
    └── Form1.Designer.cs                  # Designer-generated
```

## Development Commands

```bash
# Build
dotnet build RemuneracionQuincenal/RemuneracionQuincenal.slnx

# Run
dotnet run --project RemuneracionQuincenal/RemuneracionQuincenal.csproj

# Add NuGet packages (when implementing)
dotnet add RemuneracionQuincenal/RemuneracionQuincenal.csproj package ExcelDataReader
dotnet add RemuneracionQuincenal/RemuneracionQuincenal.csproj package DocumentFormat.OpenXml
```

## Phase Roadmap

1. **Fase 1 (Current):** Single ASE prototype — R1, R2, R4 reading + CONSOLIDADO calculation
2. **Fase 2:** Scale to all 5 ASE, add conciliation sheets, bank reports
3. **Fase 3:** Error handling, CLI mode, full logging, documentation

## Common Pitfalls

- **Dynamic row positions:** Source file rows aren't fixed — search by header values, not row numbers
- **Column 11 check:** In R2, check if header col 11 is "Especiales" before reading SERV_ESP_K
- **R4 negative values:** The reversal value is stored as negative in column 4
- **Round before write:** DetRetri column D must be integer-rounded values
- **Template formula dependencies:** Some formulas reference dynamic sub-rows — reading from source avoids this
