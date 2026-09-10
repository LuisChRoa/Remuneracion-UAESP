# Plan 18 — HU-18: Rediseño UI WinForms para usuarios administrativos (Fase 3, UX)

> **Historia:** convertir la UI técnica actual (700×580, 3 GroupBoxes planos, rutas truncadas, progreso oculto, códigos crudos en status) en una interfaz presentable para usuarios administrativos no técnicos, SIN tocar cálculo ni logging. Prioridad: CALIDAD VISUAL.
> **Rector (IRRENUNCIABLE):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 3 (cierre/entrega), §6 (trazabilidad: "Cada paso se registra"), §10 CA-4/CA-6 (fórmulas intactas, log por paso). **Invariante del proyecto:** NUNCA tocar `Core`/`Infrastructure`; solo capa `Remuneracion.WinForms`. Comportamientos existentes preservados: `Periodo.Parse`, `ProcesadorRemuneracion/Periodo`, `ArchivoFuenteLocator`, catálogo HU-14 (`MostrarErrorUx`), config Serilog HU-14/15.
> **Continuidad:** HU-01..HU-17 cerradas (UI + CLI funcionales; Q1/Q2 certificables en Capa A). Este plan NO reabre su semántica: **cero cambios de cálculo/escritura/validación; Q1/Q2 intactos por construcción.**
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-10

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. Las decisiones de juicio se dictaminan ejecutivamente en §0.3; la aprobación del plan las fija. No hay T0 de workbook en esta HU: ningún cambio toca celdas, fórmulas ni valores (solo presentación WinForms).

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (lectura directa de código):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | `Form1.Designer.cs` = 419 líneas; `ClientSize` 700×580; 3 GroupBoxes (`grpPeriodo`, `grpArchivos`, `grpEjecucion`); `txtLog` + `StatusStrip` sueltos; `MinimumSize` 650×500 | Lectura `Form1.Designer.cs:123-403` | Base del rediseño: layout actual plano sin jerarquía visual |
| V2 | `txtCarpetaFuentes/txtPlantilla/txtCarpetaSalida` = 420 px `ReadOnly`, sin `ToolTip`, sin ellipsis → rutas largas truncadas sin forma de verlas | Lectura `Form1.Designer.cs:219-285` | Justifica R2 (ToolTip con ruta completa) |
| V3 | `progressBar` 646×8 px con `Visible=false`; solo se muestra durante ejecución; sin etiqueta de % ni de ASE actual | Lectura `Form1.Designer.cs:342-347`, `Form1.cs:171-177` | Justifica R5 (progreso visible + % + ASE) |
| V4 | Status usa códigos crudos: `"Completado (salida 0)"`, `"Error {codigo} (salida N)"`, `"Cancelado (salida 5)"` | Lectura `Form1.cs:204,224,253,265` | Justifica R4 (mensajes humanos; códigos quedan en log + `UltimoCodigoSalida`) |
| V5 | `SetControlesHabilitados(bool)` existe y cubre 13 controles; es el único gate de habilitación; `btnEjecutar_Click` lo llama en inicio/fin + 2 early-returns | Lectura `Form1.cs:616-630,171,205,225,270` | R3/R5 se integran extendiéndolo, no duplicándolo |
| V6 | `UltimoCodigoSalida` (contrato HU-14/HU-15) + `MostrarErrorUx` (catálogo) + `ConfigurarSerilog` + `RunId`/`LogContext` + formatos `txtLog.AppendText($"[{HH:mm:ss}] …")` son comportamiento observable de otras HUs | Lectura `Form1.cs:28-32,285-289,632-650` | Intocables: el plan los declara invariantes con prueba de regresión |
| V7 | NO existen en WinForms: `ToolTip`, `Process.Start`, `Clipboard`, `TableLayoutPanel`, `btnLimpiar/btnAbrir/btnCopiar`, etiqueta de %/ASE/conteo-líneas | `grep` en `Remuneracion.WinForms/` (1 match: solo `MinimumSize`) | Todo R1..R6 entra como código nuevo, sin colisiones de nombres |
| V8 | Sin `Test project`, sin `.editorconfig`/`.gitattributes` decididos; estándar CRLF Windows; build `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings | `project-context.md` + planes previos | Verificación = build 0 warnings + smoke manual + regresión Q1/Q2 |

**NO verificado (y por eso NO entra):** métricas de usabilidad con usuarios reales (fuera de alcance); instalador/setup (no pedido); designer visual del IDE (el plan escribe `Designer.cs` a mano con coordenadas exactas, verificable por build + smoke).

### 0.2 Mapeo al Rector (in vs out)

**Entra (solo `Remuneracion.WinForms/`):**

| Requisito usuario | Superficie de cambio |
|---|---|
| R1 Header institucional | Nuevo `Panel` + `Label`s en `Form1.Designer.cs`, system colors |
| R2 Card layout + ellipsis/ToolTip + ventana 880×640 + Anchor/Dock + MinimumSize | Reescritura de `InitializeComponent` (controles existentes reubicados, mismos nombres) |
| R3 `btnLimpiar` nuevo | Nuevo botón + handler en `Form1.cs` + extensión de `SetControlesHabilitados` |
| R4 `btnAbrirSalida` + `btnCopiarLog` + status humano | 2 botones + handlers (`Process.Start`, `Clipboard`) + textos de `toolStripStatusLabel` |
| R5 Progreso visible + % + ASE actual + bloqueo Ejecutar/Limpiar | `progressBar` siempre visible + 2 `Label`s + updates en callback `Progress<string>` |
| R6 Log presentable (monoespaciado, autoscroll, conteo) | Propiedades de `txtLog` + helper central `AppendLogLine` (mismo formato HU-14) + `lblLineasLog` |
| R7 AI Slop detection | Checklist de verificación en U3 (§4.5) |

**Sale (EXPLÍCITO):** cualquier cambio en `Core`/`Infrastructure`; nuevos paquetes NuGet; cambios al catálogo HU-14, a Serilog, al `RunId`, a `Periodo.Parse`, a `ArchivoFuenteLocator`; cambios a flags/códigos CLI HU-15; gradientes/colores hardcodeados; nuevos formularios; commits (los hace el Ingeniero con `#commit`).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **Status humanizado, contrato intacto:** el texto visible del `StatusStrip` deja de mostrar `(salida N)` y usa mensajes humanos (§3.4). `UltimoCodigoSalida`, `Log.*` con códigos y `MostrarErrorUx` NO cambian (contrato HU-14/HU-15 para CLI y file-log). |
| G2 | **System colors únicamente:** header con `SystemColors.Highlight` / `HighlightText`; resto con `SystemColors.Control*` / `WindowText`. Cero `Color.FromArgb` / hex hardcodeado (anti-AI-slop + respeta tema Windows del equipo). |
| G3 | **TextBox no tiene ellipsis — ToolTip es el mecanismo:** WinForms `TextBox` no soporta `AutoEllipsis` (eso es de `Label`). Se mantiene `TextBox ReadOnly` (permite seleccionar/copiar la ruta) + `ToolTip` con ruta completa actualizado en `TextChanged` + doble-clic selecciona todo. Sin inventar controles de terceros. |
| G4 | **`btnLimpiar` NO borra rutas ni período:** solo `txtLog`, progreso, status y `UltimoCodigoSalida=Ok`. Borrar rutas sería destructivo para el flujo administrativo (re-ejecutar mismo período). |
| G5 | **Log con helper central, formato congelado:** todo append al log pasa por `AppendLogLine(string)` que conserva el formato HU-14 `[{HH:mm:ss}] mensaje` + autoscroll + conteo. Prohibido cambiar prefijos (el file-log y la Capa B dependen de ellos). |
| G6 | **Progreso honesto:** la barra usa los mismos hitos existentes (`maxProgreso` 8/42); solo se agrega etiqueta % (= `Value/Maximum`) y ASE actual (del mensaje de progreso del procesador, sin parseo nuevo: se muestra el texto tal cual). Sin re-mapear hitos = sin riesgo de cálculo. |
| G7 | **Regresión = workbooks Q1/Q2 5-ASE idénticos en valores visibles** (criterio HU-16/Capa A ±0.5 vs goldens `Docs/Insumos/`), porque la UI no escribe Excel: si los valores difieren, el cambio tocó lo prohibido → rollback total. |

---

## 1. PROPOSE

### 1.1 Intent

Que un usuario administrativo (no técnico) pueda ejecutar la remuneración quincenal sin ayuda: entienda en qué período está, vea sus rutas completas, sepa qué está pasando durante la ejecución y obtenga el archivo de salida y el log con un clic — con una ventana que se vea profesional y no un prototipo de ingeniería.

### 1.2 In Scope

- Re-layout de `Form1` a 880×640 con header institucional + 4 cards (`TableLayoutPanel`): Período / Rutas / Ejecución / Resultado-Log.
- 3 botones nuevos: `btnLimpiar` (obligatorio), `btnAbrirSalida`, `btnCopiarLog` + 4 labels nuevos (`lblVersionBadge`, `lblProgresoPct`, `lblAseActual`, `lblLineasLog`) + 1 `ToolTip`.
- Lógica code-behind solo de presentación (limpiar, abrir carpeta, copiar log, %/autoscroll/conteo, status humano).
- Sistema de espaciado 4 pt + `Anchor`/`Dock` correctos + `MinimumSize` 860×600 + fuente monoespaciada en log.

### 1.3 Non-Goals

Todo §0.2 "Sale": lógica de negocio, paquetes, catálogo, Serilog, CLI, goldens, commits.

### 1.4 Resultado de negocio

El Ingeniero entrega a los administrativos una app que no requiere explicación técnica (las rutas se ven, el progreso se entiende, los errores hablan en su idioma) y certifica que el rediseño no alteró ni un valor de cálculo (regresión Q1/Q2 idéntica).

---

## 2. DESIGN

### 2.1 Wireframe (880×640, `ClientSize = 880, 640`, `MinimumSize = 860, 600`)

```
┌──────────────────────────────────────────────────────────────┐
│ HEADER (Panel Dock=Top, H=64, Highlight)                     │
│  Remuneración Quincenal UAESP            [v1.0.0 badge]      │
│  Período 202607 · 2.ª Quincena  (subtitle, se actualiza)     │
├──────────────────────────────────────────────────────────────┤
│ MAIN (TableLayoutPanel Dock=Fill, 1 col × 4 filas, Pad=8)    │
│ ┌ CARD Período (H≈64) ────────────────────────────────────┐  │
│ │ Año:[____]  Mes:[________]  Quincena:[__________]        │  │
│ ├ CARD Rutas ─────────────────────────────────────────────┤  │
│ │ Fuentes:  [............................] [Seleccionar…] │  │
│ │ Plantilla:[............................] [Seleccionar…] │  │
│ │ Salida:   [............................] [Seleccionar…] │  │
│ ├ CARD Ejecución (H≈96) ──────────────────────────────────┤  │
│ │ ASE:[__________] [x] Procesar los 5 ASE                  │  │
│ │ [▶ Ejecutar] [Limpiar] [Abrir salida] [Copiar log]       │  │
│ │ [████████████░░░░░░░░]  62%   ASE 3 en curso…            │  │
│ ├ CARD Resultado-Log (Fill restante) ─────────────────────┤  │
│ │ Log de ejecución                          128 líneas [Copiar]│
│ │ ┌────────────────────────────────────────────────────┐ │  │
│ │ │ [Consolas monoespaciado, autoscroll]               │ │  │
│ │ └────────────────────────────────────────────────────┘ │  │
├──────────────────────────────────────────────────────────────┤
│ STATUS (Listo | Procesando… | Completado — archivo listo…)   │
└──────────────────────────────────────────────────────────────┘
```

### 2.2 Árbol de controles (nombres exactos; `*` = nuevo)

```
Form1 (880×640, Min 860×600, StartPosition CenterScreen)
├── pnlHeader: Panel (Dock=Top, H=64, BackColor=Highlight) *
│   ├── lblTitulo: Label ("Remuneración Quincenal UAESP", 14pt Bold, HighlightText)
│   ├── lblSubtitulo: Label ("Período …", 9pt, HighlightText) *
│   └── lblVersionBadge: Label ("v1.0.0", BorderStyle FixedSingle, HighlightText) *
├── tlpMain: TableLayoutPanel (Dock=Fill, 1×4, Padding=8, gaps 8) *
│   ├── grpPeriodo (existente, reubicado: Dock=Fill, Margin=4, Padding=8)
│   │   └── lblAnio/cmbAnio/lblMes/cmbMes/lblQuincena/cmbQuincena (sin cambios)
│   ├── grpRutas (renombrado desde grpArchivos — Text="Rutas") *
│   │   ├── lblCarpeta/txtCarpetaFuentes/btnSeleccionarCarpeta (Anchor: txt Left|Right)
│   │   ├── lblPlantilla/txtPlantilla/btnSeleccionarPlantilla
│   │   └── lblCarpetaSalida/txtCarpetaSalida/btnSeleccionarSalida
│   ├── grpEjecucion (Dock=Fill, crece a H≈110 para barra + labels)
│   │   ├── lblAse/cmbAse/chkCincoAse/btnEjecutar (sin cambios)
│   │   ├── btnLimpiar / btnAbrirSalida / btnCopiarLog *
│   │   ├── progressBar (Visible=TRUE siempre, H=18, Anchor Left|Right) *
│   │   ├── lblProgresoPct ("0 %") + lblAseActual ("—") *
│   ├── grpLog (Text="Resultado · Log", Dock=Fill) *
│   │   ├── lblLineasLog ("0 líneas") *
│   │   └── txtLog (Dock=Fill, Consolas 9pt, WordWrap=false, ScrollBars=Both)
├── statusStrip (toolStripStatusLabel Spring + toolStripVersion "v1.0.0")
├── toolTipRutas: ToolTip (AutoPopDelay=10000, ShowAlways=true) *
└── diálogos existentes (folderBrowserDialog, openFileDialogPlantilla)
```

Notas de diseño:
- **Espaciado 4 pt:** tokens XS=4 (gaps internos), S=8 (`Margin` cards + `Padding` forms), M=12 (separación header/main). Todo `Margin/Padding` múltiplo de 4; prohibido posicionar a ojo.
- **Anchoring:** `txt*` rutas `Anchor=Left|Right` (crecen con la ventana — fin del truncado estructural); `txtLog` `Dock=Fill` dentro de `grpLog`; botones de selección `Anchor=Right`; `progressBar` `Anchor=Left|Right`.
- **TabOrder:** Período (1-3) → Rutas (seleccionar 1-3) → ASE/Ejecutar/Limpiar/Abrir/Copiar → log (solo lectura, `TabStop=false`).

### 2.3 Estrategia ToolTip (G3)

- Un solo componente `toolTipRutas: ToolTip` (`AutoPopDelay=10000`, `InitialDelay=400`, `ShowAlways=true`).
- En `TextChanged` de cada `txt*` (3 handlers de 1 línea o 1 handler compartido): `toolTipRutas.SetToolTip(txt, string.IsNullOrWhiteSpace(txt.Text) ? "Sin seleccionar" : txt.Text)`.
- Doble-clic en cada `txt*` → `SelectAll()` (facilita copiar la ruta manualmente).
- Verificación con `grep`: hoy no hay `ToolTip` (V7) → sin colisiones.

### 2.4 Wiring de botones nuevos

| Botón | Handler | Lógica (solo presentación) |
|---|---|---|
| `btnLimpiar` (Text="Limpiar", en `grpEjecucion`) | `btnLimpiar_Click` | `txtLog.Clear()`; `progressBar.Value=0`; `lblProgresoPct="0 %"`; `lblAseActual="—"`; `lblLineasLog="0 líneas"`; `toolStripStatusLabel="Listo"`; `UltimoCodigoSalida=CodigosSalida.Ok`; `Log.Debug("Log limpiado por el usuario.")`. NO toca rutas/combos (G4). Habilitado solo en idle (integrado a `SetControlesHabilitados`). |
| `btnAbrirSalida` (Text="Abrir salida") | `btnAbrirSalida_Click` | Carpeta = `txtCarpetaSalida.Text`; si vacía/inexistente → `MessageBox` aviso (Warning, sin catálogo: es validación UX, no error de proceso). Si existe → `Process.Start(new ProcessStartInfo(carpeta){UseShellExecute=true})`. `using System.Diagnostics` (ya disponible en WinForms; verificar `grep` que no haya alias en conflicto). |
| `btnCopiarLog` (Text="Copiar log", en header de `grpLog` o junto a Abrir) | `btnCopiarLog_Click` | Si `txtLog.TextLength==0` → aviso; si no → `Clipboard.SetText(txtLog.Text)` + status "Log copiado al portapapeles". |
| Status humano (§3.4) | En `btnEjecutar_Click` (edición de literales) | `"Procesando…"`, `"Completado — archivo listo en {carpeta}"`, `"Cancelado por el usuario"`, `"Error {codigo} — ver guía en pantalla"` (código visible para trazabilidad, sin `(salida N)`). |

### 2.5 Estados visuales (anti-AI-slop: sin estados muertos)

| Estado | Ejecutar | Limpiar | Abrir | Copiar | Progreso |
|---|---|---|---|---|---|
| Idle sin rutas | Enabled* | Enabled | Enabled | Disabled (log vacío) | 0 %, "—" |
| Idle con log | Enabled* | Enabled | Enabled | Enabled | conserva último % |
| Ejecutando | Disabled | Disabled | Disabled | Disabled | visible, % vivo, ASE vivo |
| Error/Cancelado | Enabled | Enabled | Enabled | Enabled (hay log) | conserva valor final (no se oculta) |

`*` `btnEjecutar` mantiene su validación previa (MessageBox si faltan rutas). `SetControlesHabilitados` se extiende con los 4 botones + `lblAseActual` no editable.

---

## 3. SPEC

### 3.1 Requisitos funcionales

| ID | Requisito | Criterio de aceptación |
|---|---|---|
| RF-1 | Header institucional con título, subtítulo de período y badge de versión (system colors) | Ventana abre con `pnlHeader` visible; título legible; subtítulo refleja `PeriodoSeleccionado` al cambiar Año/Mes/Quincena; badge muestra mismo texto que `toolStripVersion`; `grep Color.FromArgb\|#rgb` en WinForms = 0 matches |
| RF-2 | Layout cards 880×640, espaciado múltiplo de 4, Anchor/Dock, MinimumSize | `ClientSize=880×640`, `MinimumSize=860×600`; redimensionar estira rutas + log sin solaparse; todos los `Margin/Padding` ∈ {4,8,12}; inspector visual: 4 cards separadas |
| RF-3 | Rutas con ToolTip de ruta completa | Hover/Doble-clic sobre ruta larga muestra/copia la ruta completa; `TextChanged` actualiza el tip; caso vacío muestra "Sin seleccionar" |
| RF-4 | `btnLimpiar` obligatorio | Clic con log lleno → log vacío, "0 líneas", barra 0 %, status "Listo", `UltimoCodigoSalida==0`; rutas y período INTACTOS; durante ejecución está Disabled |
| RF-5 | `btnAbrirSalida` | Con carpeta válida abre el Explorador en ella; sin carpeta muestra aviso y no lanza excepción |
| RF-6 | `btnCopiarLog` + conteo | Copia el contenido exacto del log al portapapeles; `lblLineasLog` = N líneas reales tras cada append y tras limpiar = "0 líneas" |
| RF-7 | Progreso visible con % y ASE actual | Durante 5-ASE la barra avanza 0→42 con % = `Value/Maximum` y `lblAseActual` muestra la fase en curso; en 1-ASE 0→8; al terminar queda en 100 % (no desaparece); Ejecutar/Limpiar Disabled durante run |
| RF-8 | Log presentable, formato HU-14 intacto | Fuente monoespaciada (Consolas 9-10 pt), `WordWrap=false`, autoscroll al fondo en cada línea, ScrollBars=Both; cada línea conserva `[{HH:mm:ss}] …`; mensajes de error conservan `ERROR [{codigo}]` |
| RF-9 | Status humano, contrato intacto | Ningún status visible contiene `(salida N)`; `UltimoCodigoSalida` y file-log conservan códigos (verificable en `remuneracion_log_*.txt` con `RunId`); `MostrarErrorUx` sin cambios |
| RF-10 | Regresión cálculo | Q1 + Q2 modo 5-ASE generan workbooks con valores visibles CONSOLIDADO dentro de ±0.5 vs goldens `Docs/Insumos/` (mismo criterio Capa A HU-16); `Core/` e `Infrastructure/` sin diff en `git status` |

### 3.2 Regresión obligatoria (bloqueante)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → 0 warnings, 0 errores (U0, U1, U2, U3).
2. Smoke UI manual (§4.5): abrir, seleccionar período, ejecutar 1 ASE si hay insumos, Limpiar, Abrir salida, Copiar log, redimensionar.
3. Regresión Q1/Q2 5-ASE: ejecutar período `202607` Q1 y Q2 con `Procesar los 5 ASE`, comparar CONSOLIDADO D9:D13/D28:D32/D47:D51/D66:D70/D85:D89/D104:D108/D109 vs goldens (tolerancia ±0.5). Si difiere → rollback (la UI tocó lo prohibido).
4. `git status --porcelain` → solo `Remuneracion.WinForms/Form1.Designer.cs`, `Remuneracion.WinForms/Form1.cs`, `plans/18*` modificados/nuevos. Cualquier diff en `Core/`/`Infrastructure/` = STOP.

---

## 4. TASKS

> Trazabilidad: cada unidad cita sus RF. Verificación por unidad: build 0 warnings + smoke incremental. Sin commits (regla del proyecto).

### U0 — Baseline verde + candado de regresión (RF-10)

1. `git status --porcelain` (anotar baseline) + `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → debe dar 0 warnings antes de tocar nada.
2. Si hay insumos Q1/Q2 disponibles: corrida 5-ASE Q1 y Q2 con la UI actual, guardar los 2 workbooks como `baseline_Q1.xlsx` / `baseline_Q2.xlsx` en temp (fuera del repo) + anotar SHA256. Si no hay insumos a mano: registrar "baseline diferida a U3" y seguir (la regresión final sigue siendo bloqueante).
3. **Verificación U0:** build 0/0 + SHAs anotados (o nota de diferimiento). Sin código tocado.

### U1 — Designer: layout header + cards + nuevos controles (RF-1, RF-2, RF-3, RF-7, RF-8)

Archivos: `Remuneracion.WinForms/Form1.Designer.cs` (único).

1. Form: `ClientSize=880×640`, `MinimumSize=860×600`, `Text` sin cambios ("Remuneración Quincenal UAESP — Fase 2" — el título institucional vive en `lblTitulo`, no en la barra).
2. Agregar `pnlHeader` (Dock=Top H=64, `BackColor=SystemColors.Highlight`) + `lblTitulo` (14 pt Bold, `ForeColor=HighlightText`) + `lblSubtitulo` (9 pt) + `lblVersionBadge` (borde fijo). Declarar campos + `Controls.Add` + `Suspend/ResumeLayout` siguiendo el patrón existente del archivo.
3. Agregar `tlpMain: TableLayoutPanel` (Dock=Fill, 1 col × 4 filas, `Padding=8`; filas: AutoSize/AutoSize/AutoSize/100 %) y mover `grpPeriodo`, `grpArchivos→Text="Rutas"`, `grpEjecucion`, nuevo `grpLog` dentro (todos `Dock=Fill`, `Margin=4`, `Padding=8`).
4. En `grpEjecucion`: `progressBar` (`Visible=true` siempre, H=18, `Anchor=Left|Right`) + `lblProgresoPct` + `lblAseActual` + `btnLimpiar` + `btnAbrirSalida` + `btnCopiarLog` (declaración + posición + `Click +=` como los botones existentes). En `grpLog`: `lblLineasLog` + mover `txtLog` (Dock=Fill, `Font=Consolas 9pt`, `WordWrap=false`, `TabStop=false`).
5. Agregar `toolTipRutas: ToolTip` (`AutoPopDelay=10000`, `ShowAlways=true`).
6. Anchors: `txt*` rutas `Left|Right`; botones `...` `Right`; `cmb*` sin cambios.
7. **Verificación U1:** build 0 warnings (el Designer roto no compila — el compilador es el primer test) + abrir la app: header visible, 4 cards, ventana 880×640, redimensión sin solapes. Botones nuevos aún sin lógica (clic = no-op hasta U2; consignarlo).

### U2 — Code-behind: Limpiar / Abrir / Copiar / progreso / status / log (RF-4, RF-5, RF-6, RF-7, RF-8, RF-9)

Archivos: `Remuneracion.WinForms/Form1.cs` (único). Prohibido tocar firmas existentes salvo `SetControlesHabilitados` (extensión) y literales de status.

1. Helper `AppendLogLine(string mensaje)` (privado): `txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {mensaje}{Environment.NewLine}")` + autoscroll (`SelectionStart=TextLength; ScrollToCaret()`) + `lblLineasLog.Text = $"{txtLog.Lines.Length} líneas"`. Migrar TODOS los `txtLog.AppendText` existentes a este helper (mismo texto, cero cambios de formato).
2. `btnLimpiar_Click` + `btnAbrirSalida_Click` (con `UseShellExecute=true`, guards) + `btnCopiarLog_Click` (guard vacío) según §2.4. `using System.Diagnostics;` solo si `grep` confirma que no hay colisión (V7 dice que no hay `Process` en el proyecto).
3. `ToolTip`: handler compartido `Ruta_TextChanged` (3 suscripciones en `Designer` o en constructor) + doble-clic `SelectAll`.
4. Progreso: en `btnEjecutar_Click`, quitar `progressBar.Visible=true/false` (siempre visible; se resetea `Value=0`, `lblProgresoPct="0 %"`, `lblAseActual="Iniciando…"`) y en el callback `Progress<string>` agregar `%` + `lblAseActual=mensaje` (texto tal cual, sin parseo). Early-returns y `finally` actualizan `lblAseActual` ("—" en idle / fase final).
5. Status humano: reemplazar los 4 literales (`"Procesando..."`, `"Completado (salida…)"`, `"Cancelado (salida…)"`, `"Error … (salida…)"`) por §2.4; NO tocar `UltimoCodigoSalida` ni `Log.*` ni `MostrarErrorUx`.
6. `SetControlesHabilitados`: agregar `btnLimpiar/btnAbrirSalida/btnCopiarLog` (+ `btnEjecutar` ya está); `lblSubtitulo` se actualiza en `InicializarPeriodo` + cambios de combos (`PeriodoSeleccionado`).
7. **Verificación U2:** build 0 warnings + smoke funcional: Limpiar (log/líneas/barra/status/código), Copiar (pegar en Notepad = idéntico), Abrir (explorador / aviso), ejecución corta con % y ASE vivos, error provocado (ruta salida == plantilla) muestra guía HU-14 + status humano + código en file-log.

### U3 — Polish anti-slop + verificación final + cierre (RF-1..RF-10)

1. Checklist AI-Slop (todo debe ser SÍ): ☐ sin gradientes/imágenes; ☐ jerarquía header>cards>log; ☐ espaciado múltiplo de 4; ☐ cero colores hardcodeados (`grep -rn "FromArgb\|Color\.[A-Z]" Remuneracion.WinForms/` = 0 salvo `SystemColors`/`System.Drawing.Point-Size` de layout); ☐ estados visuales §2.5 todos definidos (sin botón muerto); ☐ foco visible + TabOrder lógico; ☐ textos en español administrativo (sin "(salida 0)", sin tecnicismos en botones).
2. Regresión bloqueante (§3.2): build 0 warnings + Q1/Q2 5-ASE ±0.5 vs goldens/baseline + `git status` limpio de `Core/`/`Infrastructure/`.
3. Smoke de redimensión: 880×640 → 860×600 (mínimo) → maximizado: sin solapes, sin truncado estructural (rutas crecen, log crece).
4. Actualizar `toolStripVersion`/badge si el repo fija otra versión (verificar contra `.csproj`; hoy "v1.0.0" en `Designer.cs:388`).
5. **Verificación U3 (cierre):** tabla RF-1..RF-10 marcada pasa/falla + SHAs Q1/Q2 + build 0/0. Reportar al Ingeniero para aprobación. Sin commit.

---

## 5. Architecture Validation Certificate

| Principio | Veredicto | Justificación |
|---|---|---|
| SRP | ✅ | Cada handler nuevo hace una sola cosa de presentación (limpiar / abrir / copiar / tip). Cálculo y logging quedan en sus clases existentes. |
| OCP | ✅ | Se extiende `SetControlesHabilitados` y `InitializeComponent`; no se modifican `Procesador*`, `Locator` ni catálogo. |
| DIP | ✅ | `Form1` sigue dependiendo de `IProcesadorRemuneracion/IProcesadorPeriodo`_ctor-inyectados; lo nuevo (`Process`, `Clipboard`, `ToolTip`) son servicios de plataforma, no abstracciones de dominio. |
| Best practices WinForms | ✅ | `Suspend/ResumeLayout`, `Anchor/Dock`, `UseShellExecute=true`, guards antes de IO, `BeginUpdate/EndUpdate` existentes intactos, system colors (respeta tema/alto contraste). |
| Performance | ✅ | Costo O(1) por línea de log (conteo + scroll); ToolTip en `TextChanged` (evento raro); sin timers ni layouts anidados profundos (1 `TableLayoutPanel`, sin `DataGridView`). Ejecución sigue en `Task.Run` — la UI no bloquea. |
| Seguridad | ✅ | `Process.Start` solo con ruta elegida por el usuario vía diálogo (no input libre); `Clipboard` solo bajo clic explícito; sin secretos ni paths hardcodeados. |

---

## 6. Riesgos y mitigaciones

| Riesgo | Prob. | Mitigación |
|---|---|---|
| El Designer escrito a mano no compila (coordenadas/controles) | Media | U1 verifica con build; el compilador es el test; cambios pequeños y build tras cada card |
| Status humano rompe alguna aserción HU-14/HU-15 | Baja | G1: códigos intactos en `UltimoCodigoSalida` + file-log; `grep "(salida "` confirma 0 literales visibles; tests no existen en UI (sin suite que romper) |
| `lblAseActual=mensaje` muestra texto técnico del procesador | Media | Aceptado por G6 (honesto > inventado); el mensaje ya es el que hoy va al log; si el Ingeniero lo rechaza en smoke, se acota a "Procesando ASE N…" con parseo mínimo en U3 |
| Regresión Q1/Q2 difiere (tocó cálculo sin querer) | Muy baja | Candado §3.2.4 (`git status` sin `Core/`/`Infra`); rollback = `git checkout -- Remuneracion.WinForms/` |
| `Clipboard` falla en sesión RDP / `Process.Start` sin explorador | Baja | Guards + `try/catch` con aviso (no catálogo); el flujo principal (Ejecutar) no depende de ellos |

---

*Plan generado por sdd-planner — Fuente Única de Verdad para HU-18. Implementar solo con aprobación del Ingeniero. Sin código implementado en este documento.*
