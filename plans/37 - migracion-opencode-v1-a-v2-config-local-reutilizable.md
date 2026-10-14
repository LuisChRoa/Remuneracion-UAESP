# Plan 37 — Migración OpenCode config local v1 → v2 (reutilizable)

> SDD PROPOSE → DESIGN → SPEC → TASKS. Solo config `.opencode/`. Cero código de producción.

- **Fecha:** 2026-10-13
- **Estado:** PROPOSED (plan, no implementado)
- **Alcance:** `.opencode/opencode.jsonc` + `.opencode/agents/*.md` + verificación `AGENTS.md` / `project-context.md`
- **Reutilizable para:** N proyectos con config local opencode v1
- **Síntoma que motiva:** en v2 solo se invoca `general`, nunca los custom (`implementer`, `code-reviewer`, etc.)
- **Causa conocida (doc oficial v2):** renombres `agent`→`agents`, `prompt`→`system`, `permission`→`permissions[{action,resource,effect}]`, `bash`→`shell`, `task`→`subagent`; `mode: primary|subagent|all`; `hidden` saca de descubrimiento; `subagent_depth` default 1; tool `subagent` no expone catálogo de IDs al modelo (bugs #36761 / #35238).

---

## PROPOSE

### Objetivo

Recuperar delegación a subagentes custom en OpenCode v2, con una plantilla de port v1→v2 reutilizable en N repos, verificada por `opencode agent list` + invocación real `subagent(agent="<custom>")` sin fallback a `general`.

### Alcance (in)

1. Port de `.opencode/opencode.jsonc` (forma v1) a forma v2.
2. Port de frontmatter de los 6 `agents/*.md`: `remuneracion-architect` (principal) + 5 subagentes (`sdd-planner`, `implementer`, `code-reviewer`, `debug-agent`, `ui-ux-specialist`).
3. Reglas de `hidden`, `mode`, `subagent_depth`, `default_agent`, modelos, permisos, skills y MCPs.
4. Checklist de validación + rollback por tarea.
5. Checklist/script mental de replicación cruzada para otros repos.

### No-objetivos (out)

- No tocar `Remuneracion.Core|Infrastructure|WinForms`, ni `*.slnx|csproj`.
- No crear test project, no cambiar paquetes NuGet.
- No migrar skills globales (`~/.config/opencode/skills/`) a locales salvo decisión explícita.
- No cambiar prompts/contenido de agentes, solo frontmatter + config.
- No hacer commit/push sin `#commit` / `#push` explícito (regla del principal).

### Éxito medible

- `opencode agent list` muestra los 6 customs.
- `subagent(agent="implementer"|"code-reviewer"|...)` resuelve al custom; cero fallback a `general`/`explore`.
- `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` sigue en 0 warnings (no se tocó código, es guarda).
- `plans/` ocupa el 37 sin colisionar con 01..36 existentes.

---

## DESIGN

### Decisión D1 — Formato destino: `opencode.jsonc` v2 (no `opencode.json`)

Mantener JSONC (comentarios permitidos). Estructura v2 mínima:

```jsonc
{
  "$schema": "https://opencode.ai/config.json",
  "default_agent": "{{PRIMARY_AGENT}}",
  "agents": {
    "{{PRIMARY_AGENT}}": {
      "mode": "primary",
      "model": "<provider>/<model-sin-variant>",
      "system": "{file:agents/{{PRIMARY_AGENT}}.md}",
      "permissions": [
        { "action": "subagent", "resource": "<cada-subagente>", "effect": "allow" },
        { "action": "subagent", "resource": "explore", "effect": "allow" },
        { "action": "subagent", "resource": "general", "effect": "allow" }
      ],
      "subagent_depth": 2
    }
  }
}
```

Por qué: `agent` (singular) es v1 e ignorado en v2; `agents` (plural) es lo que v2 descubre. `default_agent` se conserva pero solo funciona si el agente referenciado es visible con `mode: primary|all`.

### Decisión D2 — Tabla de mapeo v1 → v2 exhaustiva (clave por clave)

| # | v1 (actual) | v2 (destino) | Regla de port | Riesgo si se omite |
|---|-------------|--------------|---------------|--------------------|
| 1 | `agent: { <id>: {...} }` (top-level singular) | `agents: { <id>: {...} }` (plural) | Renombrar clave. Todo lo demás cuelga de aquí. | Config ignorada; solo `general` disponible. |
| 2 | `prompt: "{file:...}"` | `system: "{file:...}"` | Renombrar. Sintaxis `{file:...}` se conserva. | Agente sin system prompt; comportamiento genérico. |
| 3 | `permission: { task: {...} }` | `permissions: [{action,resource,effect}]` | `task`→`action: "subagent"`. Cada entrada `"id": "allow"` → `{action:"subagent", resource:"<id>", effect:"allow"}`. `"*": "allow"` → prohibido como wildcard ciego; expandir a lista explícita. | Permisos v1 ignorados → principal no puede delegar. |
| 4 | frontmatter `permission: { edit: allow/deny, bash: {...} }` | frontmatter `permissions` o `permission` v2 según doc vigente del repo | `edit` se conserva como está salvo doc diga `tools.write`. `bash: { "dotnet build*": allow }` → `shell: { "dotnet build*": allow }`. `webfetch: deny` → `webfetch: deny` o entrada `permissions`. Verificar con `opencode agent list --verbose` tras port. | Build bloqueado o webfetch abierto. |
| 5 | `bash:` (en permission) | `shell:` | Renombrar. Glob patterns (`dotnet build*`) se conservan. | Reglas shell no aplican. |
| 6 | `task:` (action name) | `subagent:` | Renombrar en permisos del principal. | Delegación denegada. |
| 7 | `mode: "primary"` (solo principal) | `mode: "primary"` | Sin cambio de valor, pero ahora es excluyente: primary NO aparece como subagente invocable. | Si se pone `all` por error, el principal se auto-delega. |
| 8 | `mode: "subagent"` (5 customs) | `mode: "subagent"` | Mantener. NO cambiar a `all` salvo que el custom deba usarse también como default interactivo. | `all` expone customs en TUI; ruido. |
| 9 | `hidden: true` (5 subagentes) | `hidden: false` o eliminar clave | **Decisión: quitar `hidden:true`.** En v2 `hidden` saca de descubrimiento y contribuye al síntoma (invisibles para `subagent`). Si se quiere ocultar del TUI pero mantener invocables, validar en v2 si `hidden` + invocación por ID sigue funcionando; por defecto asumir que NO. | Customs invisibles → fallback a `general`. Causa #1 del síntoma junto con `agent` vs `agents`. |
| 10 | (ausente) `subagent_depth` | `subagent_depth: 2` en el principal | Default v2 = 1 (el subagente ya no puede delegar). Poner 2 solo en el principal para permitir `principal → sdd-planner → implementer` si el flujo lo exige; dejar 1 (omitir) en subagentes. | Con depth 1, delegación en cadena falla silenciosamente. |
| 11 | `default_agent: "remuneracion-architect"` | igual | Conservar. Debe apuntar a un agente con `mode: primary|all` y visible (`hidden` falso/ausente). | Default cae a `general`. |
| 12 | `model: "opencode/muse-spark-1.3-contributor-free"` | igual `provider/model`, variant con `#` | **Estrategia modelos:** mantener `provider/model` tal cual; si hay sufijo de variant (`-free`, `-thinking`), mover a formato `provider/model#variant` solo si la doc v2 del entorno lo exige. Validar cada ID con lista de modelos disponibles; si uno no existe, fijar fallback explícito y registrar en plan. No mezclar `opencode-go/` con `opencode/` sin verificar que el registry los resuelve. | Modelo inválido → agente no carga. |
| 13 | `description:` en frontmatter | igual | Conservar verbatim. Es lo que el router usa para elegir subagente. | Sin description, el router no discrimina y manda a `general`. |
| 14 | skills globales `~/.config/opencode/skills/` | sin cambio por defecto | **Estrategia skills:** no copiar a `.opencode/skills/` (hoy vacío a propósito). Solo crear symlink/copia local si un repo necesita pin de versión. `project-context.md` sigue inyectado junto con skills al delegar (vía system prompt del principal). | Duplicar skills diverge versiones; no hacerlo sin motivo. |
| 15 | MCPs (no hay locales) | declarar solo si se usan | Si un subagente necesita MCP, añadir bloque `mcp:` explícito en v2; no asumir herencia del principal. | Permisos MCP heredados no aplican en v2. |
| 16 | `"$schema"` | igual | Conservar. | Sin schema, sin autocompletado ni validación. |

### Decisión D3 — `hidden: true → false`

Quitar `hidden: true` en los 5 subagentes. Justificación: el síntoma es invisibilidad; `hidden` en v2 es la forma explícita de invisibilizar. Costo: aparecen en TUI — aceptable frente a no ser invocables. Alternativa descartada: mantener `hidden:true` + invocación directa por ID — solo válida si se demuestra con `subagent(agent="implementer")` exitoso en validación; si falla, aplicar D3 sin excepción.

### Decisión D4 — `subagent_depth`

Principal `subagent_depth: 2`. Subagentes: omitir (default 1). Evita cascadas infinitas `implementer → debug-agent → implementer`. Si un flujo exige 3 niveles, subir a 3 solo para ese plan y revertir después.

### Decisión D5 — `default_agent` visible primary

`default_agent` debe referenciar al principal con `mode: primary` y sin `hidden`. Verificación: arrancar sesión y confirmar que el agente activo es `remuneracion-architect`, no `general`.

### Decisión D6 — Modelos

No cambiar de familia durante la migración. Solo normalizar sintaxis si v2 lo exige (`#variant`). Tabla de pin actual → destino:

| Agente | v1 actual | v2 destino (propuesto) |
|--------|-----------|------------------------|
| remuneracion-architect | `opencode/muse-spark-1.3-contributor-free` | igual, o `opencode/muse-spark-1.3#contributor-free` si el registry lo exige |
| sdd-planner / implementer / code-reviewer / debug-agent / ui-ux-specialist | `opencode-go/...` o `opencode/...` (ver frontmatter de cada archivo) | mismo provider/model, normalizar `#` igual que arriba |

Cada ID se valida contra el registry antes de cerrar el port (tarea T6).

### Decisión D7 — Skills y project-context

- `.opencode/skills/` permanece vacío (skills globales mandan).
- `.opencode/project-context.md` intacto; el principal lo inyecta al delegar (ya es convención del repo).
- `AGENTS.md` raíz: solo añadir nota de v2 (nombres de claves) si el port cambia algo que el lector deba saber; no reescribir.

### Decisión D8 — MCPs

Sin MCPs locales hoy → no añadir bloque `mcp:`. Si un repo reutilizador sí tiene, mapear `mcp_<server>:` v1 a formato v2 y listar servidores por agente.

---

## SPEC

### Plantilla genérica reutilizable (variables)

```jsonc
{
  "$schema": "https://opencode.ai/config.json",
  "default_agent": "{{PRIMARY_AGENT}}", // ej. "remuneracion-architect"
  "agents": {
    "{{PRIMARY_AGENT}}": {
      "mode": "primary",
      "model": "{{PROVIDER}}/{{MODEL}}",
      "system": "{file:agents/{{PRIMARY_AGENT}}.md}",
      "permissions": [
        // {{SUBAGENTS[]}} expandido uno por uno, SIN wildcard "*"
        { "action": "subagent", "resource": "{{SUBAGENT_1}}", "effect": "allow" },
        { "action": "subagent", "resource": "{{SUBAGENT_2}}", "effect": "allow" },
        { "action": "subagent", "resource": "explore", "effect": "allow" },
        { "action": "subagent", "resource": "general", "effect": "allow" }
      ],
      "subagent_depth": 2
    }
    // NOTA: los subagentes se definen SOLO en agents/*.md (frontmatter), no duplicarlos aquí
    // salvo que la doc v2 del entorno exija definición inline.
  }
}
```

Frontmatter v2 por subagente (`agents/{{SUBAGENT}}.md`):

```md
---
mode: subagent
model: {{PROVIDER}}/{{MODEL}}
# hidden: omitir (visible). Solo poner hidden: true si se demostró invocable oculto.
permissions:
  shell:
    "dotnet build*": allow
    "dotnet test*": allow
  edit: allow
  webfetch: deny
description: "{{DESCRIPCION_VERBATIM_V1}}"
---
```

### Checklist de port por archivo

- [ ] `P1` `.opencode/opencode.jsonc`: `agent`→`agents`, `prompt`→`system`, `permission.task.*`→`permissions[subagent/*]`, expandir `"*"`, añadir `subagent_depth: 2` al principal.
- [ ] `P2` `agents/remuneracion-architect.md`: confirmar `mode: primary` (o añadir si es el principal sin frontmatter mode), `hidden` ausente, `model` válido, `description` presente.
- [ ] `P3..P7` cada `agents/{sdd-planner,implementer,code-reviewer,debug-agent,ui-ux-specialist}.md`: `mode: subagent`, quitar `hidden: true`, `bash`→`shell`, validar `model`, conservar `description` verbatim.
- [ ] `P8` `AGENTS.md`: nota mínima v2 (opcional, 3 líneas).
- [ ] `P9` `.opencode/project-context.md`: intacto (verificar diff vacío).
- [ ] `P10` `.opencode/skills/`: sigue vacío (sin copias globales).

### Ejemplo mínimo antes/después (config)

Antes (v1, actual):

```jsonc
{ "default_agent": "remuneracion-architect",
  "agent": { "remuneracion-architect": {
    "mode": "primary", "model": "opencode/muse-spark-1.3-contributor-free",
    "prompt": "{file:agents/remuneracion-architect.md}",
    "permission": { "task": { "*": "allow", "sdd-planner": "allow" } } } } }
```

Después (v2):

```jsonc
{ "default_agent": "remuneracion-architect",
  "agents": { "remuneracion-architect": {
    "mode": "primary", "model": "opencode/muse-spark-1.3-contributor-free",
    "system": "{file:agents/remuneracion-architect.md}",
    "permissions": [
      { "action": "subagent", "resource": "sdd-planner", "effect": "allow" },
      { "action": "subagent", "resource": "implementer", "effect": "allow" },
      { "action": "subagent", "resource": "code-reviewer", "effect": "allow" },
      { "action": "subagent", "resource": "debug-agent", "effect": "allow" },
      { "action": "subagent", "resource": "ui-ux-specialist", "effect": "allow" },
      { "action": "subagent", "resource": "explore", "effect": "allow" },
      { "action": "subagent", "resource": "general", "effect": "allow" }
    ],
    "subagent_depth": 2 } } }
```

### Ejemplo completo frontmatter subagente

Antes (v1):

```md
---
mode: subagent
hidden: true
model: opencode-go/algun-modelo
permission:
  edit: allow
  bash:
    "dotnet build*": allow
  webfetch: deny
description: "Implementa desde plan aprobado..."
---
```

Después (v2):

```md
---
mode: subagent
model: opencode-go/algun-modelo
permissions:
  edit: allow
  shell:
    "dotnet build*": allow
  webfetch: deny
description: "Implementa desde plan aprobado..."
---
```

(Cambio: fuera `hidden`, `bash`→`shell`, `permission`→`permissions` según doc vigente; `description` intacta.)

### Criterios de aceptación verificables

1. `opencode agent list` (o equivalente del CLI instalado) muestra `remuneracion-architect` + los 5 customs. Evidencia: captura/salida pegada en el PR o en `plans/37-T0-Evidencia.md`.
2. Invocación real `subagent(agent="implementer")` con tarea trivial (“responde tu ID y modelo”) resuelve al custom, no a `general`. Repetir para `code-reviewer` y `sdd-planner` (muestra mínima 3/5).
3. `default_agent` activo al abrir sesión = `remuneracion-architect`.
4. `git diff -- .opencode/` solo toca `opencode.jsonc` + frontmatters; `project-context.md` sin diff; cero `*.cs` tocados.
5. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` = 0 warnings (guarda de no-regresión).
6. Rollback por tarea disponible (ver TASKS) y probado al menos en `opencode.jsonc` (`git stash` / `git checkout --`).

---

## TASKS (atómicas, ordenadas)

> Cada tarea: archivo → cambio exacto → verificación (comando) → rollback.

- [ ] **T1 — Backup.** `git status --short` limpio o con cambios anotados; `cp .opencode/opencode.jsonc C:\Users\lechaparro.ASEINGES\AppData\Local\Temp\opencode\opencode.jsonc.v1.bak` + `git diff -- .opencode/ > C:\Users\lechaparro.ASEINGES\AppData\Local\Temp\opencode\pre37.diff`. Verificación: existen ambos archivos. Rollback: n/a (es el respaldo).
- [ ] **T2 — Snapshot frontmatters.** `grep -H "^mode:\|^hidden:\|^model:\|^prompt\|^permission\|^bash" .opencode/agents/*.md` (o `rg`) y guardar salida en `plans/37-T0-Evidencia-A.md` (anexo, no parte del plan). Verificación: anexo creado. Rollback: borrar anexo.
- [ ] **T3 — Port `opencode.jsonc` (estructura).** Archivo: `.opencode/opencode.jsonc`. Cambios: `agent`→`agents`; `prompt`→`system`; `permission.task`→`permissions[]` expandido (5 customs + explore + general, sin `"*"`); añadir `"subagent_depth": 2` al principal. Verificación: `python3 -c "import json,sys; json.load(open('.opencode/opencode.jsonc'))"` o validador JSONC del editor pasa. Rollback: `git checkout -- .opencode/opencode.jsonc`.
- [ ] **T4 — Frontmatter principal.** Archivo: `.opencode/agents/remuneracion-architect.md`. Cambios: asegurar `mode: primary`, sin `hidden`, `model` válido, `description` presente (añadir una línea si falta). Verificación: `head -20` muestra frontmatter correcto. Rollback: `git checkout -- .opencode/agents/remuneracion-architect.md`.
- [ ] **T5 — Frontmatter 5 subagentes (uno por commit/hunk).** Archivos: `sdd-planner.md`, `implementer.md`, `code-reviewer.md`, `debug-agent.md`, `ui-ux-specialist.md`. Cambios por archivo: quitar `hidden: true`; `bash`→`shell`; `permission`→`permissions` (según doc); conservar `model`+`description`. Verificación por archivo: diff muestra solo esas 3 líneas. Rollback por archivo: `git checkout -- <archivo>`.
- [ ] **T6 — Validar modelos.** Resolver cada `model:` contra el registry del entorno (`opencode models list | grep <id>` o equivalente). Si un ID no existe, fijar fallback y anotar en `plans/37-T0-Evidencia.md`. Verificación: los 6 IDs resuelven. Rollback: revertir el pin del modelo en el archivo afectado.
- [ ] **T7 — Validación cruzada de descubrimiento.** Comandos: `opencode agent list` (debe listar 6 customs) + invocación `subagent(agent="implementer")` tarea trivial + `subagent(agent="code-reviewer")` + `subagent(agent="sdd-planner")`. Evidencia en `plans/37-T0-Evidencia.md`. Criterio: 3/3 resuelven al custom, 0 fallback a `general`. Rollback: n/a (solo lectura); si falla, volver a T3/T5.
- [ ] **T8 — Guarda de no-regresión.** `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → 0 warnings. `git diff --stat -- .opencode/` + `git diff --name-only -- *.cs *.csproj` (este último debe estar vacío). Rollback: n/a.
- [ ] **T9 — Nota mínima en AGENTS.md (opcional).** Añadir ≤5 líneas: claves v2 vigentes + `hidden` + `subagent_depth`. Verificación: `git diff -- AGENTS.md` ≤10 líneas. Rollback: `git checkout -- AGENTS.md`.
- [ ] **T10 — Replicación cruzada (otros repos).** Checklist/script mental: (1) `ls .opencode/agents/ + cat .opencode/opencode.jsonc`; (2) aplicar misma tabla D2; (3) `diff -u <bak-v1> <nuevo-v2>` archivado; (4) repetir T7 en ese repo; (5) registrar desviaciones (modelos distintos, MCPs, skills locales) como follow-ups. Verificación: checklist firmado por repo en su propio `plans/NN`. Rollback: restaurar `.bak` del paso 1.

Orden de commits sugerido: T3 → T4 → T5 (5 micro-commits) → T6 → T9. T1/T2/T7/T8/T10 no commitean (evidencia local o planes).

---

## Architecture Validation Certificate

- **Contract-first:** el plan no cambia contratos de dominio (`Core/Interfaces`, modelos, fórmulas CONSOLIDADO). Solo config de orquestación.
- **Blast radius:** confinado a `.opencode/` + nota en `AGENTS.md`. Cero producción. Cada tarea tiene rollback de un comando.
- **Reversibilidad:** `git checkout -- .opencode/` + restauración del `.bak` revierte el 100%.
- **Observabilidad:** T7/T8 exigen evidencia (`37-T0-Evidencia.md`) y guarda de build.
- **Reutilización:** SPEC §Plantilla + T10 hacen el port replicable sin re-diseñar.
- **Deuda conocida:** si la doc v2 del entorno difiere en `permissions` vs `permission` para frontmatter, T5 debe ajustarse a la doc instalada (la tabla D2 fila 4 lo marca como punto de verificación, no como verdad absoluta).

## Riesgos

| # | Riesgo | Probabilidad | Impacto | Mitigación (tarea) |
|---|--------|--------------|---------|--------------------|
| R1 | Pérdida de permisos: `permissions[]` mal expandido deja al principal sin delegación | Media | Alto | T3 sin wildcard + T7 con 3 invocaciones reales |
| R2 | Modelos inválidos (`opencode-go/...` no resuelve en v2) | Media | Alto | T6 valida los 6 IDs antes de cerrar |
| R3 | `hidden` reintroducido por hábito | Media | Alto | D3 + T5 diff de 3 líneas revisado |
| R4 | `subagent_depth=1` rompe cadenas `planner→implementer` | Baja | Medio | D4 (`depth: 2` solo principal) + T7 en cadena si aplica |
| R5 | Skills path: copiar globales a local y divergir | Baja | Medio | D7: `.opencode/skills/` vacío + P10 en checklist |
| R6 | `subagent` no expone catálogo (bugs #36761/#35238) persiste aun con config perfecta | Baja | Alto | T7 es el test que lo distingue de un error de config; si T7 falla con config correcta, escalar a issue upstream con evidencia |
| R7 | Colisión `plans/37*` con convención `01..36` | Baja | Bajo | Nombre exacto `37 - migracion-...md`; T2/T7 usan sufijo `-T0-` para anexos |

## Follow-ups (no parte de este plan)

- F1: pin de versiones de skills globales si algún repo necesita reproducibilidad.
- F2: guía MCP por agente cuando aparezca el primer MCP local (hoy no hay).
- F3: si v2 estabiliza `hidden` invocable, re-evaluar D3 para ocultar customs del TUI.
- F4: replicar en 1 repo piloto distinto antes de rodar a N (T10).

---

*Fin plan 37. Implementación solo vía `implementer` con plan aprobado; commits solo con `#commit`/`#push` explícito.*
