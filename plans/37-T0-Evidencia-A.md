# Plan 37 — T0 Evidencia-A: snapshot pre-port (v1)

Fecha: 2026-10-13. Fuente única de verdad: `plans/37 - migracion-opencode-v1-a-v2-config-local-reutilizable.md` (T2).

## A1. Frontmatters pre-port (Select-String `^(name:|mode:|hidden:|model:|description:|permission|...)`)

```text
code-reviewer.md:2: name: code-reviewer
code-reviewer.md:3: description: "Subagente de remuneracion-architect para revisión de calidad y spec compliance contra el plan SDD de plans/. Prioridad del proyecto: calidad senior."
code-reviewer.md:4: mode: subagent
code-reviewer.md:5: model: opencode-go/glm-5.3-flash
code-reviewer.md:6: hidden: true
code-reviewer.md:7: permission:
code-reviewer.md:8:   edit: deny
code-reviewer.md:9:   bash:
code-reviewer.md:16:   webfetch: deny
debug-agent.md:2: name: debug-agent
debug-agent.md:3: description: "Subagente de remuneracion-architect para análisis sistemático de causa raíz. NO arregla sin investigar primero. Prioridad: calidad senior."
debug-agent.md:4: mode: subagent
debug-agent.md:5: model: opencode-go/deepseek-v4.1-flash
debug-agent.md:6: hidden: true
debug-agent.md:7: permission:
debug-agent.md:8:   edit: deny
debug-agent.md:9:   bash:
debug-agent.md:16:   webfetch: deny
implementer.md:2: name: implementer
implementer.md:3: description: "Subagente de remuneracion-architect para implementación siguiendo el plan SDD aprobado. Contract-first, traza cada cambio a un requisito del plan de plans/. TDD cuando aplique."
implementer.md:4: mode: subagent
implementer.md:5: model: opencode-go/deepseek-v4.1-flash
implementer.md:6: hidden: true
implementer.md:7: permission:
implementer.md:8:   edit: allow
implementer.md:9:   bash:
implementer.md:15:   webfetch: deny
remuneracion-architect.md:2: name: remuneracion-architect
remuneracion-architect.md:3: description: "Agente principal para Remuneración Quincenal UAESP — Orquestador, mentor técnico senior y guardián de calidad especializado en .NET 10 (Desktop WinForms) + Excel. Delega a subagentes, enseña y verifica el flujo SDD completo."
sdd-planner.md:2: name: sdd-planner
sdd-planner.md:3: description: "Subagente de remuneracion-architect para planificación SDD formal. Ejecuta PROPOSE → DESIGN → SPEC → TASKS y genera documento física en plans/. Prioridad del proyecto: calidad senior."
sdd-planner.md:4: mode: subagent
sdd-planner.md:5: model: opencode/muse-spark-1.3-contributor-free
sdd-planner.md:6: hidden: true
sdd-planner.md:7: permission:
sdd-planner.md:8:   edit: allow
sdd-planner.md:9:   bash:
sdd-planner.md:15:   webfetch: deny
ui-ux-specialist.md:2: name: ui-ux-specialist
ui-ux-specialist.md:3: description: "Subagente de remuneracion-architect para diseño premium, UX y accesibilidad en WinForms (.NET 10). Traduce el Visual Design Intent del plan SDD a código. Detecta AI Slop."
ui-ux-specialist.md:4: mode: subagent
ui-ux-specialist.md:5: model: opencode/muse-spark-1.3-contributor-free
ui-ux-specialist.md:6: hidden: true
ui-ux-specialist.md:7: permission:
ui-ux-specialist.md:8:   edit: allow
ui-ux-specialist.md:9:   bash:
ui-ux-specialist.md:15:   webfetch: deny
```

Conclusión: principal SIN `mode:`/`model:` (solo `name`+`description`); 5 subagentes con `mode: subagent` + `hidden: true` + `permission:`/`bash:` (forma v1).

## A2. `opencode.jsonc` pre-port: `agent` (singular), `prompt`, `permission.task` con `"*": "allow"` — ver `Temp/opencode/opencode.jsonc.v1.bak`.

## A3. `opencode debug agents` pre-port (IDs únicos, global+local mezclados)

```text
big-pickle, build, code-reviewer, compaction, debug-agent, deepseek-v4.1-flash,
explore, general, gentle-ai-explore, gentle-ai-verify, gentle-ai-worker,
gentle-orchestrator, glm-5.3-flash, implementer, jd-fix-agent, jd-judge-a,
jd-judge-b, jira-assistant, muse-spark-1.3-contributor-free, n8n-Architect,
plan, remuneracion-architect, review-readability, review-refuter,
review-reliability, review-resilience, review-risk, review-validator,
sdd-planner, summary, title, ui-ux-specialist
```

Los 6 customs del repo YA resuelven pre-port (el `hidden:true` no los saca de `debug agents`; solo del TUI).

## A4. Modelos pre-port — los 3 resuelven en `opencode models` (T6)

- `opencode/muse-spark-1.3-contributor-free` — exact match.
- `opencode-go/deepseek-v4.1-flash` — exact match.
- `opencode-go/glm-5.3-flash` — exact match.

## A5. Line endings pre-port: todos CRLF puros (loneLF=0) — `opencode.jsonc` + 6 `agents/*.md`.

## A6. T1 backup

- `git status --short` pre-port: `M Remuneracion.Core/Errors/*` (2 archivos), untracked: `.engram/`, `Docs/Plantilla estandarizada/`, `Docs/Prueba Agosto-2/Resultado/Resultado6/`, `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs`, 2 tests integration, `plans/36*`, `plans/37*`. Cambios anotados, no limpios — se procede (T1 permite "limpio o con cambios anotados").
- Backup: `Temp/opencode/opencode.jsonc.v1.bak` (638 B) + `Temp/opencode/pre37.diff` (0 B — `.opencode/` sin diff vs HEAD).
