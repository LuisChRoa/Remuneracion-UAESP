# Plan 37 — T0 Evidencia: validación post-port (v2)

Fecha: 2026-10-13. CLI: `opencode v2.0.28`. Comandos usados: `opencode reload` + `opencode debug agents`
(nota: en v2.0.28 NO existe subcomando `opencode agent list`; el equivalente es `opencode debug agents`).

## B1. Agentes post-port (6/6 resuelven, `hidden=False` en los 6)

```text
code-reviewer          mode=subagent hidden=False model=opencode-go/glm-5.3-flash
debug-agent            mode=subagent hidden=False model=opencode-go/deepseek-v4.1-flash
implementer            mode=subagent hidden=False model=opencode-go/deepseek-v4.1-flash
remuneracion-architect mode=primary  hidden=False model=opencode/muse-spark-1.3-contributor-free
sdd-planner            mode=subagent hidden=False model=opencode/muse-spark-1.3-contributor-free
ui-ux-specialist       mode=subagent hidden=False model=opencode/muse-spark-1.3-contributor-free
```

## B2. Permisos efectivos del principal (merge global + local)

Incluyen los 7 del port local, verificados en el JSON de `debug agents`:

```text
subagent sdd-planner allow
subagent implementer allow
subagent code-reviewer allow
subagent debug-agent allow
subagent ui-ux-specialist allow
subagent explore allow
subagent general allow
```

System prompt del principal carga desde `{file:agents/remuneracion-architect.md}` (verificado: el `system`
parseado contiene el markdown del principal). `subagent_depth` no se expone en `debug agents`; consta en
`.opencode/opencode.jsonc` (`"subagent_depth": 2`) y el JSONC valida con `ConvertFrom-Json`.

## B3. T6 modelos — los 3 IDs resuelven en `opencode models` (match exacto)

- `opencode/muse-spark-1.3-contributor-free` (principal, sdd-planner, ui-ux-specialist)
- `opencode-go/deepseek-v4.1-flash` (implementer, debug-agent)
- `opencode-go/glm-5.3-flash` (code-reviewer)

Sin normalización `#variant`: innecesaria, los IDs v1 resuelven tal cual en v2.0.28.

## B4. T8 no-regresión

- `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → `Build succeeded. 0 Warning(s) 0 Error(s)`.
- `git diff --stat -- .opencode/` → solo los 7 archivos del port (ver T3/T4/T5).
- `.opencode/project-context.md` y `.opencode/skills/` → sin diff (intactos / vacío).
- CRLF puro (loneLF=0) en los 7 archivos `.opencode/` + anexos de evidencia.

## B5. Desviaciones y asuntos (trazabilidad)

1. **D-COLISION-GLOBAL (resuelta en el port):** existe `~/.config/opencode/agents/ui-ux-specialist.md`
   global con `hidden: true` y otra descripción (web/n8n). v2 hace merge: las claves ausentes en el local
   se heredan del global → tras quitar `hidden` del local, el efectivo seguía `hidden=True`.
   Fix: `hidden: false` explícito SOLO en `.opencode/agents/ui-ux-specialist.md` (D3 lo permite;
   post-fix efectivo `hidden=False`, descripción y modelo = locales). Los otros 4 subagentes no colisionan
   (sin homónimo global) y quedan sin `hidden` según SPEC. Follow-up para el Inge: decidir si se renombra
   o retira el global (fuera de alcance del plan: prohíbe tocar config global sin decisión explícita).
2. **D-CS-PREEXISTENTE (no causada por este port):** `git diff --name-only -- *.cs *.csproj` NO está vacío:
   `M Remuneracion.Core/Errors/CatalogoErrores.cs` y `M Remuneracion.Core/Errors/CodigoError.cs`
   (+11/-1) ya estaban modificados ANTES del port (consta en `37-T0-Evidencia-A.md` §A6). Este port no tocó
   ningún `.cs`/`.csproj`; el build da 0 warnings con esos cambios presentes.
3. **D-INVOCACION-PENDIENTE (para el Inge):** la invocación real `subagent(agent="implementer")` con tarea
   trivial (criterio 2 del plan, 3/5 customs) no puede ejecutarla el implementer (sin herramienta subagent
   anidada). Comando sugerido en sesión principal:
   `opencode run "Responde solo tu ID y modelo" --agent implementer` (repetir con `code-reviewer` y
   `sdd-planner`; éxito = responde el custom, sin fallback a `general`).
   En esta CLI `opencode run --help` debe confirmar el flag exacto (`--agent` u otro) antes de correrlo.
4. **D-AGENTS-MD-ASCII (resuelta):** la nota T9 en `AGENTS.md` se escribe en ASCII puro (sin backticks) porque
   el archivo raíz no es UTF-8 y PowerShell interpreta backticks en strings doble-comentados; un primer intento
   con escapes dejó bytes corruptos (BEL 0x07) y se revirtió por truncado byte-exacto + re-anexado. Diff final:
   solo 6 líneas añadidas, cero modificaciones.

## B6. T10 replicación cruzada

El checklist del plan (§TASKS T10) queda vigente + 2 pasos nuevos derivados de este port:
(0) `opencode debug agents` (no `agent list`) como comando de verificación en v2.0.28;
(6) tras el port, comparar cada `hidden` efectivo contra su homónimo en `~/.config/opencode/agents/`
y fijar `hidden: false` explícito donde haya colisión (caso D-COLISION-GLOBAL).
