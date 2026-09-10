# Plan 19 — HU-19: UI premium light colapsable (Fase 3, UX)

> **Historia:** elevar la UI HU-18 (880×640, header full-Highlight, 4 GroupBoxes, botones de igual peso, log siempre visible) a una interfaz premium light para usuarios administrativos: fondo claro, cards blancas con borde fino, jerarquía de botones, y Card Resultado **colapsada por defecto** (estado humano + una línea de resumen; el log técnico solo bajo "Ver detalle técnico"). Sin tocar cálculo, logging, CLI ni paquetes.
> **Rector (IRRENUNCIABLE):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 3 (cierre/entrega), §6 (trazabilidad: "Cada paso se registra"), §10 CA-4/CA-6 (fórmulas intactas, log por paso). **Invariante del proyecto:** NUNCA tocar `Core`/`Infrastructure`/`Cli`; solo capa `Remuneracion.WinForms`. Comportamientos existentes preservados: `Periodo.Parse`, `ProcesadorRemuneracion/Periodo`, `ArchivoFuenteLocator`, catálogo HU-14 (`CatalogoErrores`/`MostrarErrorUx`), contrato `UltimoCodigoSalida` HU-14/15, `ConfigurarSerilog` + `RunId`/`LogContext`, formato de línea `txtLog` `[{HH:mm:ss}]`.
> **Continuidad:** HU-01..HU-18 cerradas (HEAD = `fcc0802` "Rediseño UI WinForms admin según HU-18"). Este plan construye **encima de los controles HU-18** (mismos nombres salvo renombre documentado en §2.2) y **absorbe S-1** (cleanup). **W-1 excluido** (§0.2).
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Inge.
> **Fecha:** 2026-09-10

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Inge. Las decisiones de juicio se dictaminan ejecutivamente en §0.3; la aprobación del plan las fija. No hay T0 de workbook: ningún cambio toca celdas, fórmulas ni valores (solo presentación WinForms).

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (lectura directa de código + git + grep):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | Árbol limpio; HU-18 es HEAD `fcc0802` (commit, NO "uncommitted" como asumía el encargo: `git status` = `nothing to commit, working tree clean`) | `git status` + `git log --oneline -8` + `git show --stat HEAD` | Baseline U0 parte de árbol limpio; cualquier diff futuro es 100 % HU-19 |
| V2 | `Form1.Designer.cs` = 658 líneas; `ClientSize` 880×640; `MinimumSize` 860×600; header `pnlHeader` Dock=Top H=64 con `BackColor=Highlight` + `lblTitulo` 14pt Bold + `lblSubtitulo` ("Período …") + `lblVersionBadge`; `tlpMain` 1 col × 4 filas (3×AutoSize + 100 %), `Padding=8` | Lectura `Form1.Designer.cs:164-233` | Base exacta del rediseño: se reutilizan `pnlHeader`, `tlpMain`, combos, `toolTipRutas`, `statusStrip`, diálogos |
| V3 | 4 `GroupBox`: `grpPeriodo`, `grpArchivos` (Text="Rutas"), `grpEjecucion`, `grpLog` (Text="Resultado · Log"); `progressBar` siempre visible + `lblProgresoPct` + `lblAseActual`; `txtLog` Consolas 9pt + `lblLineasLog`; `btnLimpiar/btnAbrirSalida/btnCopiarLog` existen con `UseVisualStyleBackColor=true` | Lectura `Form1.Designer.cs:34-81, 423-596` | Gana HU-19: GroupBox→Panel, jerarquía de botones vía `FlatStyle`, log colapsable |
| V4 | S-1 confirmado: `lblCarpetaSalida/txtCarpetaSalida/btnSeleccionarSalida` se inicializan con `= new …()` **en la declaración de campo** (`Designer.cs:55-57`) mientras el resto se crea con `new` dentro de `InitializeComponent` (funciona —los inicializadores de campo corren antes— pero es estilo inconsistente y frágil ante regeneración del Designer) | Lectura `Form1.Designer.cs:55-57` vs `91-152` | S-1 se absorbe en U1: mover los 3 `new` a `InitializeComponent`, cero cambio de comportamiento |
| V5 | `Form1.cs` = 799 líneas; invariantes intactos: `UltimoCodigoSalida` (D3), `AppendLogLine` formato `[{HH:mm:ss}]` + autoscroll + conteo, `Ruta_TextChanged`/`Ruta_DoubleClick`, `SetControlesHabilitados`, `ConfigurarSerilog` (rolling diario, template canónico), `RunId`/`Periodo`/`Modo` en `LogContext`, `MostrarErrorUx` por catálogo, status humanos HU-18 | Lectura `Form1.cs` completa | Intocables salvo lo listado en §0.2 "Entra"; el plan los declara invariantes con prueba de regresión |
| V6 | Cero colisiones para lo nuevo: `grep` por `pnlCard\|cardPeriodo\|chkVerDetalle\|lnkVerDetalle\|lblEstadoHumano\|lblResumenUnaLinea` = 0 matches; `FlatStyle` = 0 matches (jerarquía por documentar desde cero); `BorderStyle` solo en `lblVersionBadge`; `Color\.`/`FromArgb` = 0 matches (cero colores hardcodeados hoy) | `grep` en `Remuneracion.WinForms/` | Todo lo premium entra como código nuevo sin renombres forzados |
| V7 | `Program.cs`: DI por constructor (`IProcesadorRemuneracion`, `IProcesadorPeriodo`, `ArchivoFuenteLocator` + oráculo HU-13); `Remuneracion.WinForms.csproj`: `net10.0-windows`, Serilog 4.4.0 + Sinks.File 7.0.0, WinExe | Lectura `Program.cs` + `.csproj` | Sin cambios: NO nuevos NuGet, NO cambios de composición |
| V8 | Sin test project; build `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` con 0 warnings (convención); CRLF Windows; sin commit/push (los hace el Inge con `#commit`) | `project-context.md` + `AGENTS.md` | Verificación = build 0 warnings + smoke manual + regresión Q1/Q2 |

**NO verificado (y por eso NO entra):** métricas de usabilidad con usuarios reales; instalador/setup; designer visual del IDE (el plan escribe `Designer.cs` a mano con coordenadas exactas, verificable por build + smoke); inspección pixel-perfect de render (se valida en smoke U3, no en este documento).

### 0.2 Mapeo al Rector (in vs out)

**Entra (solo `Remuneracion.WinForms/Form1.Designer.cs` + `Form1.cs`):**

| Requisito usuario | Superficie de cambio |
|---|---|
| Sistema visual light (tokens §2.3, Segoe UI 9-10pt, spacing 4pt, gaps 8-12, cards Padding 12-16) | `InitializeComponent`: colores `SystemColors`, fuentes, `Margin`/`Padding`; ventana 960×680 / mínima 940×660 |
| Header slim (título + subtítulo live de período + badge versión derecha, UNA línea de acento en vez de fill Highlight) | Reestilado de `pnlHeader` + nuevo `pnlAccent` (H=2-3, `Highlight`); `lblTitulo/lblSubtitulo/lblVersionBadge` reutilizados |
| Cards como `Panel` blancos con borde fino + título 10pt semibold con barra de acento (fin del look sunken de `GroupBox`) | `grpPeriodo/grpArchivos/grpEjecucion/grpLog` → `pnlCardPeriodo/pnlCardRutas/pnlCardEjecucion/pnlCardResultado` (renombre documentado §2.2); títulos `lblCardTitulo*` + `pnlAccentBar*` |
| Card Período con labels encima de los campos | Re-layout interno con `TableLayoutPanel` 3 cols (mismos `cmbAnio/cmbMes/cmbQuincena`) |
| Card Rutas: 3 filas `TextBox ReadOnly` + ToolTip full-path + botones compactos `...` | `txtCarpetaFuentes/txtPlantilla/txtCarpetaSalida` reutilizados; `btnSeleccionar*` reutilizados con `Text="..."`, ancho ~32, + ToolTip "Examinar…" |
| Card Ejecución: fila ASE + check 5-ASE, fila de botones con jerarquía, fila progreso % + ASE actual | `cmbAse/chkCincoAse/progressBar/lblProgresoPct/lblAseActual` reutilizados; `btnEjecutar` primario solid; `btnLimpiar/btnAbrirSalida/btnCopiarLog` secundarios flat-outline (§2.5) |
| Card Resultado colapsada por defecto (estado humano + resumen 1 línea; `CheckBox` "Ver detalle técnico" expande log monoespaciado con conteo + Copiar) | Nuevos `lblEstadoHumano`, `lblResumenUnaLinea`, `chkVerDetalle`, `pnlDetalleTecnico` (contiene `txtLog` + `lblLineasLog` movidos); handler `chkVerDetalle_CheckedChanged` |
| S-1: mover los 3 `new` de campos (Designer 55-57) a `InitializeComponent` | U1, mecánico, cero comportamiento |
| TabOrder lógico; Anchor/Dock responsive | `TabIndex` + `Anchor`/`Dock` en Designer |

**Sale (EXPLÍCITO):** cualquier cambio en `Core`/`Infrastructure`/`Cli`; nuevos paquetes NuGet; cambios al catálogo HU-14, a Serilog, al `RunId`, a `Periodo.Parse`/`PeriodoSeleccionado`, a `ArchivoFuenteLocator`, a flags/códigos CLI HU-15; gradientes/imágenes/colores hex hardcodeados; nuevos formularios; **W-1** (tweak de modelo en `.opencode/agents/implementer.md` —nota: HEAD ya tocó ese archivo 1 línea en HU-18; HU-19 NO lo toca); commits.

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **Colapso sin redimensionar el form:** la fila 4 de `tlpMain` sigue `Percent 100 %`; colapsar oculta solo `pnlDetalleTecnico` (la card muestra resumen arriba). Sin resize dinámico = sin riesgo de flicker/layout. Si al Inge le molesta el aire sobrante en smoke, U3 contempla alternativa acotada (fila a `AutoSize` + `MinimumSize` menor) como follow-up, NO en este plan. |
| G2 | **Tokens `SystemColors` únicamente, documentados en §2.3:** cero `Color.FromArgb`/hex (anti-AI-slop + respeta tema/alto contraste de Windows). `grep Color\.FromArgb` debe dar 0. |
| G3 | **Renombre GroupBox→Panel SÍ se hace** (excepción a "mismos nombres"): `grp*` → `pnlCard*` porque el tipo cambia (`GroupBox`→`Panel`) y el nombre debe decir la verdad. Es mecánico y grep-verificable; `Form1.cs` no referencia `grp*` (verificado: solo Designer los usa) → riesgo ~0. |
| G4 | **Botones `...` conservan nombres** (`btnSeleccionarCarpeta/Plantilla/Salida`): cambia `Text` a `"..."` + `Size` ~32×28 + ToolTip "Examinar…". `Form1.cs` no toca su `Text` → cero lógica afectada. |
| G5 | **Primario solid sin violar "no random hex":** `btnEjecutar.FlatStyle=Flat`, `BackColor=Highlight`, `ForeColor=HighlightText`, `UseVisualStyleBackColor=false`. Secundarios: `FlatStyle=Flat`, `BackColor=Window`, `ForeColor=WindowText`, `FlatAppearance.BorderColor=ControlDark`, `UseVisualStyleBackColor=false`. Todo `SystemColors` → respeta tema. |
| G6 | **Resumen humano honesto, sin parseo nuevo:** `lblEstadoHumano` refleja el mismo `toolStripStatusLabel` (idle "Listo", "Procesando…", "Completado — archivo listo…", "Error {codigo} — ver guía"); `lblResumenUnaLinea` = `Período {X} · {Modo} · {GranTotal o archivo}` construido con datos que el handler YA tiene (`PeriodoSeleccionado`, `modoTexto`, `rutaSalida`, `granTotal` 5-ASE). Sin agregar cálculo. |
| G7 | **Log intacto bajo el colapso:** `AppendLogLine`, formato `[{HH:mm:ss}]`, autoscroll, conteo y `btnCopiarLog` NO cambian; solo su contenedor se oculta. El file-log Serilog ni se entera. |
| G8 | **Regresión = workbooks Q1/Q2 5-ASE idénticos en valores visibles** (criterio Capa A ±0.5 vs goldens `Docs/Insumos/`): la UI no escribe Excel; si difieren, se tocó lo prohibido → rollback total. |

---

## 1. PROPOSE

### 1.1 Intent

Que un usuario administrativo abra la app y **entienda todo sin ayuda**: ventana clara y liviana, cada grupo con título legible, rutas visibles con acceso al path completo en un clic, un botón principal obvio (Ejecutar) y secundarios discretos, progreso que se lee en lenguaje humano, y un resultado que **por defecto le habla en su idioma** ("Completado — archivo listo en…") sin exponerle el log técnico salvo que lo pida ("Ver detalle técnico").

### 1.2 In Scope

- Reestilado completo de `Form1` a 960×680 (mínima 940×660): fondo light, header slim con línea de acento, 4 cards `Panel` blancas con títulos 10pt semibold + barra de acento.
- Jerarquía de botones (primario/secondarios) vía `FlatStyle` + `UseVisualStyleBackColor`, documentada y grep-verificable.
- Card Resultado colapsable: resumen humano + 1 línea; detalle técnico (log monoespaciado, autoscroll, conteo, Copiar) oculto por defecto tras `CheckBox`.
- Re-layouts internos: Período con labels arriba; Rutas con botones `...`; Ejecución con filas ASE / botones / progreso.
- TabOrder lógico, Anchor/Dock responsive, ToolTips, empty states definidos (§2.6).
- Absorción de S-1 (cleanup Designer 55-57).

### 1.3 Non-Goals

Todo §0.2 "Sale": lógica de negocio, paquetes, catálogo, Serilog, CLI, goldens, W-1, commits. Tampoco: temas oscuros, personalización de colores por usuario, persistencia de "detalle expandido" entre sesiones (siempre arranca colapsado — decisión, no olvido).

### 1.4 Resultado de negocio

El Inge entrega a los administrativos una app que no requiere explicación técnica (el log deja de asustar por defecto, los errores hablan en humano, el progreso se entiende) con apariencia profesional clara — y certifica que el rediseño no alteró ni un valor de cálculo (regresión Q1/Q2 idéntica) ni el contrato observável HU-14/15.

---

## 2. DESIGN

### 2.1 Wireframe (960×680, `ClientSize = 960, 680`, `MinimumSize = 940, 660`)

Estado COLAPSADO (default al abrir):

```
┌────────────────────────────────────────────────────────────────┐
│ HEADER slim (Panel, BackColor=Window, H=56)                     │
│  Remuneración Quincenal UAESP                  [v1.0.0 badge]   │
│  Período 202607 · 2.ª Quincena                                 │
│ ────────── accent line (Panel H=2, Highlight, Dock=Bottom) ────│
├────────────────────────────────────────────────────────────────┤
│ MAIN (TableLayoutPanel Dock=Fill, 1 col × 4 filas, Pad=8/12)    │
│ ┌ CARD Período ──────────────────────────────────────────────┐  │
│ │ ┃ Año        Mes          Quincena                         │  │
│ │ ┃ [____]    [________]   [__________]                     │  │
│ ├ CARD Rutas ────────────────────────────────────────────────┤  │
│ │ ┃ Carpeta fuentes                                          │  │
│ │ ┃ [.......................................] [...]          │  │
│ │ ┃ Plantilla / Carpeta salida (igual)                       │  │
│ ├ CARD Ejecución ────────────────────────────────────────────┤  │
│ │ ┃ ASE: [__________]  [x] Procesar los 5 ASE                │  │
│ │ ┃ [▶ Ejecutar]  [Limpiar] [Abrir salida] [Copiar log]      │  │
│ │ ┃   ↑ primario solid    ↑↑↑ secundarios flat-outline       │  │
│ │ ┃ [████████░░░░░░]  62%   ASE 3 en curso…                  │  │
│ ├ CARD Resultado (COLAPSADA) ────────────────────────────────┤  │
│ │ ┃ Estado: Completado — archivo listo en …                  │  │
│ │ ┃ Período 202607-2 · 5 ASE · GranTotal = 1.234.567         │  │
│ │ ┃ [ ] Ver detalle técnico                                  │  │
│ ├──────────────────────────────────────────────────────────────┤
│ STATUS (Listo | Procesando… | Completado — archivo listo…)     │
└────────────────────────────────────────────────────────────────┘
┃ = accent bar (Panel W=4, Highlight) junto al título de cada card
```

Estado EXPANDIDO (`[x] Ver detalle técnico`):

```
│ ├ CARD Resultado (EXPANDIDA) ────────────────────────────────┤  │
│ │ ┃ Estado + resumen (igual que colapsada)                   │  │
│ │ ┃ [x] Ver detalle técnico                    128 líneas    │  │
│ │ ┃ ┌────────────────────────────────────────────────────┐ │  │
│ │ ┃ │ [Consolas 9pt, autoscroll, WordWrap=false]         │ │  │
│ │ ┃ │ [HH:mm:ss] …                                       │ │  │
│ │ ┃ └────────────────────────────────────────────────────┘ │  │
```

### 2.2 Árbol de controles (nombres exactos; `*` = nuevo, `~` = renombrado G3)

```
Form1 (ClientSize 960×680, MinimumSize 940×660, BackColor=Control, StartPosition CenterScreen)
├── pnlHeader: Panel (Dock=Top, H=56, BackColor=Window)            [reestilado: antes H=64 Highlight]
│   ├── lblTitulo: Label ("Remuneración Quincenal UAESP", Segoe UI 11pt Semibold, WindowText)
│   ├── lblSubtitulo: Label ("Período …", 9pt, GrayText)           [lógica ActualizarSubtitulo intacta]
│   ├── lblVersionBadge: Label ("v1.0.0", BorderStyle FixedSingle, GrayText)
│   └── pnlAccent: Panel (Dock=Bottom, H=2, BackColor=Highlight) *
├── tlpMain: TableLayoutPanel (Dock=Fill, 1×4, Padding=12, gaps 8-12)
│   ├── pnlCardPeriodo: Panel (BackColor=Window, BorderStyle=FixedSingle, Padding=12, Margin=4) ~grpPeriodo
│   │   ├── pnlAccentBarPeriodo: Panel (W=4, Dock=Left, Highlight) *
│   │   ├── lblCardTituloPeriodo: Label ("Período", 10pt Semibold) *
│   │   └── tlpPeriodo: TableLayoutPanel (3 cols) *
│   │       ├── lblAnio / cmbAnio | lblMes / cmbMes | lblQuincena / cmbQuincena  [labels ARRIBA de cada combo]
│   ├── pnlCardRutas: Panel (igual estilo card) ~grpArchivos
│   │   ├── pnlAccentBarRutas * + lblCardTituloRutas ("Rutas") *
│   │   ├── lblCarpeta + txtCarpetaFuentes (Anchor L|R, ReadOnly) + btnSeleccionarCarpeta (Text="...", 32×28)
│   │   ├── lblPlantilla + txtPlantilla + btnSeleccionarPlantilla (idem)
│   │   └── lblCarpetaSalida + txtCarpetaSalida + btnSeleccionarSalida (idem)
│   ├── pnlCardEjecucion: Panel (igual estilo card) ~grpEjecucion
│   │   ├── pnlAccentBarEjecucion * + lblCardTituloEjecucion ("Ejecución") *
│   │   ├── fila ASE: lblAse / cmbAse / chkCincoAse (sin cambios de lógica)
│   │   ├── fila botones: btnEjecutar (primario §2.5) / btnLimpiar / btnAbrirSalida / btnCopiarLog (secundarios)
│   │   └── fila progreso: progressBar (Anchor L|R) + lblProgresoPct + lblAseActual
│   └── pnlCardResultado: Panel (igual estilo card, Dock=Fill) ~grpLog
│       ├── pnlAccentBarResultado * + lblCardTituloResultado ("Resultado") *
│       ├── lblEstadoHumano: Label (9-10pt, WindowText, AutoEllipsis) *        [espejo humano del status]
│       ├── lblResumenUnaLinea: Label (9pt, GrayText, AutoEllipsis) *          [una línea: período·modo·total]
│       ├── chkVerDetalle: CheckBox (Text="Ver detalle técnico") *            [Checked=false por defecto]
│       └── pnlDetalleTecnico: Panel (Dock=Fill, Visible=false) *
│           ├── lblLineasLog ("0 líneas", Anchor Top|Right)                    [movido]
│           └── txtLog (Dock=Fill, Consolas 9pt, WordWrap=false, ScrollBars=Both, TabStop=false) [movido]
├── statusStrip (sin cambios: toolStripStatusLabel Spring + toolStripVersion)
├── toolTipRutas: ToolTip (existente; + tips "Examinar…" en los 3 botones ...)
└── diálogos existentes (folderBrowserDialog, openFileDialogPlantilla)
```

Notas:

- `Form1.cs` NO referencia `grp*` (verificado por lectura completa: solo Designer los nombra) → el renombre G3 no arrastra code-behind.
- Orden de `Controls.Add` + `Suspend/ResumeLayout` sigue el patrón existente del archivo.
- `TabOrder` (§2.7): Período (cmbAnio→cmbMes→cmbQuincena) → Rutas (3 botones `...`) → Ejecución (cmbAse→chkCincoAse→btnEjecutar→btnLimpiar→btnAbrirSalida→btnCopiarLog) → Resultado (chkVerDetalle; `txtLog.TabStop=false`).
- `Anchor/Dock`: `txt*` rutas `Left|Right`; botones `...` `Right`; `progressBar` `Left|Right`; `lblAseActual` `Right` + `AutoEllipsis`; cards `Dock=Fill` en `tlpMain`; `txtLog` `Dock=Fill` en `pnlDetalleTecnico`.

### 2.3 Tokens visuales (única fuente; todo `SystemColors`, cero hex)

| Token | Valor | Uso |
|---|---|---|
| T-BG | `SystemColors.Control` | Fondo ventana + `tlpMain` (smoke claro del SO) |
| T-CARD | `SystemColors.Window` | Fondo de las 4 cards + header slim + botones secundarios |
| T-BORDER | `SystemColors.ControlDark` | Borde fino cards (`BorderStyle.FixedSingle` usa el color del SO; si se requiere línea custom, `FlatAppearance.BorderColor = ControlDark`) |
| T-TEXT | `SystemColors.WindowText` | Texto principal, títulos de card (10pt semibold), botones secundarios |
| T-MUTED | `SystemColors.GrayText` | Subtítulo header, `lblResumenUnaLinea`, badge versión |
| T-ACCENT | `SystemColors.Highlight` | `pnlAccent` header, 4 `pnlAccentBar*`, fondo `btnEjecutar` |
| T-ACCENT-TEXT | `SystemColors.HighlightText` | Texto de `btnEjecutar` |
| T-TITLE | Segoe UI 10pt Semibold | Los 4 `lblCardTitulo*` |
| T-BODY | Segoe UI 9pt | Resto de labels/botones/combos |
| T-HEADER-TITLE | Segoe UI 11pt Semibold | `lblTitulo` |
| T-LOG | Consolas 9pt | `txtLog` (heredado HU-18) |
| T-XS / T-S / T-M | 4 / 8 / 12 | Gaps internos / `Margin` cards + `Padding` forms / separación header-main; `Padding` cards 12-16; prohibido posicionar a ojo |

Verificación: `grep -rn "Color\.FromArgb\|Color\.FromName\|#\[0-9A-Fa-f\]" Remuneracion.WinForms/` = 0 matches (aplican `System.Drawing.Point/Size` de layout, que NO cuentan como color).

### 2.4 Estados colapsado vs expandido + empty states

| Aspecto | Colapsado (default, `chkVerDetalle.Checked=false`) | Expandido (`Checked=true`) |
|---|---|---|
| `pnlDetalleTecnico.Visible` | `false` (`txtLog` + `lblLineasLog` ocultos, fuera de TabOrder) | `true` |
| `lblEstadoHumano` | Siempre visible: espejo de `toolStripStatusLabel` | Idem |
| `lblResumenUnaLinea` | Siempre visible (1 línea, `AutoEllipsis`) | Idem |
| Arranque app | Colapsado SIEMPRE (aunque la sesión anterior quedó expandida — G §1.3) | Solo por clic explícito |
| Durante ejecución | Puede expandirse sin afectar el run (append sigue funcionando oculto) | Autoscroll/conteo en vivo como hoy |

Empty states (anti-AI-slop: ningún panel vacío sin explicación):

| Zona | Empty state |
|---|---|
| `lblEstadoHumano` inicial | "Listo — elija período y rutas y pulse Ejecutar." |
| `lblResumenUnaLinea` inicial | "Aún no hay ejecución en esta sesión." |
| `lblResumenUnaLinea` tras éxito 1-ASE | `Período {P} · ASE {n-nombre} · salida = {nombreArchivo}` |
| `lblResumenUnaLinea` tras éxito 5-ASE | `Período {P} · 5 ASE · GranTotal ≈ {granTotal:0.##}` |
| `lblResumenUnaLinea` tras error/cancelado | `Período {P} · no completado ({estado humano corto}) — ver guía en pantalla` |
| Rutas vacías | ToolTip "Sin seleccionar" (heredado HU-18) |

### 2.5 Jerarquía de botones (`FlatStyle` + `UseVisualStyleBackColor`)

| Botón | `FlatStyle` | `BackColor` | `ForeColor` | `UseVisualStyleBackColor` | Notas |
|---|---|---|---|---|---|
| `btnEjecutar` (primario) | `Flat` | `Highlight` | `HighlightText` | `false` | Solid; `FlatAppearance.BorderSize=0`; texto `"▶ Ejecutar"` (heredado); ancho ~140 |
| `btnLimpiar/btnAbrirSalida/btnCopiarLog` (secundarios) | `Flat` | `Window` | `WindowText` | `false` | Outline: `FlatAppearance.BorderColor=ControlDark`, `BorderSize=1`; hover default del SO |
| `btnSeleccionar*` (`...` compactos) | `Standard` (default) | default | default | `true` | Sin jerarquía: son utilitarios de fila; `Size≈32×28`, ToolTip "Examinar…" |

Hovers/foco: los deja el SO (respeta tema/alto contraste); prohibido owner-draw.

### 2.6 Estrategia ToolTip (hereda G3 de HU-18, se extiende)

- `toolTipRutas` existente (`AutoPopDelay=10000`, `InitialDelay=400`, `ShowAlways=true`): sin cambios de config.
- `Ruta_TextChanged` sigue actualizando el tip de los 3 `txt*` (código intacto; solo cambian de padre a `pnlCardRutas`).
- Nuevo: `toolTipRutas.SetToolTip(btnSeleccionar*, "Examinar…")` para los 3 botones compactos (el `"..."` solo se entiende con tip).
- `Ruta_DoubleClick` (`SelectAll`) intacto.

### 2.7 Wiring de collapse + resumen humano

| Elemento | Handler / punto de integración | Lógica (solo presentación) |
|---|---|---|
| `chkVerDetalle_CheckedChanged` | Nuevo en `Form1.cs` | `pnlDetalleTecnico.Visible = chkVerDetalle.Checked; tlpMain.PerformLayout();` Si se expande con log vacío, `lblLineasLog` muestra "0 líneas" (ya lo hace). |
| `lblEstadoHumano` | Helper `ActualizarResumenHumano()` invocado donde hoy se escribe `toolStripStatusLabel.Text` (mismos 6-7 sitios: idle, "Procesando…", completado, cancelado, error plantilla, error catch, "Log copiado…") | `lblEstadoHumano.Text = toolStripStatusLabel.Text;` (espejo literal — imposible que diverjan). |
| `lblResumenUnaLinea` | Mismo helper, con parámetros que el caller YA tiene | Idle: empty state; tras éxito 1-ASE/5-ASE: formatos §2.4 (usa `PeriodoSeleccionado`, `modoTexto`, `rutaSalida`/`granTotal` existentes); error: mensaje corto humano. |
| `SetControlesHabilitados` | Extensión (única firma que se toca) | Agrega `chkVerDetalle.Enabled = habilitados;` (durante run no se puede colapsar/expandir a medias; al terminar se rehabilita). `pnlDetalleTecnico` conserva su visibilidad al deshabilitar. |
| `btnLimpiar_Click` | 2 líneas nuevas | Además de lo HU-18: `chkVerDetalle.Checked = false; ActualizarResumenHumano();` (vuelve al estado colapsado inicial). |

---

## 3. SPEC

### 3.1 Criterios por card

**Card Período:**

| ID | Criterio de aceptación |
|---|---|
| AC-PER-01 | Labels "Año/Mes/Quincena" ENCIMA de cada combo (layout 3 cols), fuente T-BODY; card blanca con título 10pt semibold + barra de acento |
| AC-PER-02 | Cambiar Año/Mes/Quincena actualiza `lblSubtitulo` del header (`Período {X} · {quincena}`) — `PeriodoSeleccionado` y `Periodo.Parse` intactos |

**Card Rutas:**

| ID | Criterio de aceptación |
|---|---|
| AC-RUT-01 | 3 filas `TextBox ReadOnly` con `Anchor Left|Right` (crecen con la ventana) + botones compactos `Text="..."` `Size≈32×28` anclados a la derecha |
| AC-RUT-02 | Hover sobre cada ruta muestra el path completo; ruta vacía muestra "Sin seleccionar"; cada `...` tiene tip "Examinar…"; doble-clic selecciona todo (heredado HU-18, verificado en smoke) |
| AC-RUT-03 | Ningún path visible queda truncado sin remedio: o cabe (ventana 960) o se revela vía ToolTip/doble-clic |

**Card Ejecución:**

| ID | Criterio de aceptación |
|---|---|
| AC-EJE-01 | `btnEjecutar` visualmente primario (fondo acento + texto claro, `Flat/BorderSize=0`, `UseVisualStyleBackColor=false`); `Limpiar/Abrir salida/Copiar log` secundarios outline (`Flat/BorderColor=ControlDark`); anchos: Ejecutar > secundarios; ventana NO muestra 4 botones de igual peso |
| AC-EJE-02 | Fila ASE intacta: `cmbAse` + `chkCincoAse` ("Procesar los 5 ASE") con la misma lógica de habilitación; fila progreso con barra + `%` (=`Value/Maximum`) + ASE/fase actual en texto tal cual del procesador |
| AC-EJE-03 | Durante ejecución: Ejecutar/Limpiar/Abrir/Copiar + `chkVerDetalle` Disabled; al terminar: 100 % visible (la barra NO desaparece) |

**Card Resultado (colapsable):**

| ID | Criterio de aceptación |
|---|---|
| AC-RES-01 (colapso default) | Al abrir la app: `chkVerDetalle.Checked=false`, `pnlDetalleTecnico.Visible=false`, `txtLog` oculto; visibles SOLO `lblEstadoHumano` (empty state) + `lblResumenUnaLinea` ("Aún no hay ejecución…") + el check "Ver detalle técnico" |
| AC-RES-02 | Marcar el check expande el log (monoespaciado Consolas, `WordWrap=false`, autoscroll, `lblLineasLog` con N real, `btnCopiarLog` copia el contenido exacto); desmarcar lo oculta sin perder contenido |
| AC-RES-03 | Tras ejecución exitosa el resumen de 1 línea sigue el formato §2.4 (período·modo·total/archivo) SIN códigos técnicos (`(salida N)`, `ERR-*`, `RunId` quedan solo en file-log + `UltimoCodigoSalida`) |
| AC-RES-04 | `btnLimpiar` restaura el estado colapsado inicial (check desmarcado, empty states, "0 líneas", `UltimoCodigoSalida==Ok`; rutas y período intactos) |

** IA-slop premium (toda la ventana):**

| ID | Criterio de aceptación |
|---|---|
| AC-SLOP-01 | Cero `GroupBox` (todos → `Panel`): `grep "GroupBox" Form1.Designer.cs` = 0 matches |
| AC-SLOP-02 | Cero fondo gris full-window: ventana `BackColor=Control`, cards `Window` (contraste card/fondo visible); cero colores hardcodeados (`grep FromArgb/FromName/#hex` = 0) |
| AC-SLOP-03 | Cero gradientes/imágenes/owner-draw: `grep "LinearGradient\|DrawItem\|OnPaint" Remuneracion.WinForms/` = 0 |
| AC-SLOP-04 | Todos los `Margin/Padding` ∈ {4, 8, 12, 16}; gaps entre cards 8-12; `Padding` cards 12-16 (revisión por lectura del Designer) |
| AC-SLOP-05 | Textos 100 % español administrativo; ningún status/resumen visible contiene `(salida N)`, `ERR-`, `RunId`, `GUID` |

### 3.2 Regresión obligatoria (bloqueante)

| ID | Criterio |
|---|---|
| AC-REG-01 | Q1 + Q2 modo 5-ASE generan workbooks con valores visibles CONSOLIDADO (D9:D13/D28:D32/D47:D51/D66:D70/D85:D89/D104:D108/D109) dentro de **±0.5** vs goldens `Docs/Insumos/` (mismo criterio Capa A HU-16). Si difiere → rollback (se tocó lo prohibido). |
| AC-BLD-01 | `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → 0 warnings, 0 errores (U0, U1, U2, U3). |
| AC-GIT-01 | `git status --porcelain` → solo `Remuneracion.WinForms/Form1.Designer.cs`, `Remuneracion.WinForms/Form1.cs`, `plans/19*`. Cualquier diff en `Core/`/`Infrastructure/`/`Cli/`/`.csproj`/paquetes = STOP. Sin commit (lo hace el Inge). |

---

## 4. TASKS

> Trazabilidad: cada unidad cita sus AC. Verificación por unidad: build 0 warnings + smoke incremental. Sin commits (regla del proyecto). Orden estricto: U0 → U1 (solo Designer) → U2 (solo code-behind) → U3 (polish + verificación).

### U0 — Baseline verde + candado de regresión (AC-REG-01, AC-BLD-01, AC-GIT-01)

1. `git status --porcelain` (debe estar limpio: HEAD `fcc0802`) + `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → 0 warnings antes de tocar nada.
2. Si hay insumos Q1/Q2 a mano: corrida 5-ASE Q1 y Q2 con la UI actual, guardar los 2 workbooks como `baseline_Q1.xlsx` / `baseline_Q2.xlsx` en `temp/` (fuera del repo) + anotar SHA256. Si no hay insumos: registrar "baseline diferida a U3" y seguir (la regresión final sigue bloqueante).
3. **Verificación U0:** build 0/0 + SHAs anotados (o nota de diferimiento). Sin código tocado.

### U1 — Designer only: sistema visual + cards + colapso (AC-PER-01, AC-RUT-01, AC-EJE-01, AC-RES-01, AC-SLOP-01..04)

Archivo: `Remuneracion.WinForms/Form1.Designer.cs` (ÚNICO; prohibido tocar `Form1.cs` en U1).

1. Form: `ClientSize=960×680`, `MinimumSize=940×660`, `BackColor=SystemColors.Control`. S-1: mover los 3 `new` de campos (líneas 55-57) a `InitializeComponent` y dejar las declaraciones como el resto (mecánico, sin cambios de nombres de esos 3).
2. `pnlHeader`: `H=56`, `BackColor=Window`; `lblTitulo` 11pt semibold `WindowText` (quitar `BackColor=Highlight`); `lblSubtitulo` 9pt `GrayText`; `lblVersionBadge` `GrayText`; nuevo `pnlAccent` (`Dock=Bottom`, `H=2`, `Highlight`).
3. Renombre G3 `grp*`→`pnlCard*` (`GroupBox`→`Panel`, `BackColor=Window`, `BorderStyle=FixedSingle`, `Margin=4`, `Padding=12..16`, `Dock=Fill`); por card: `pnlAccentBar*` (`W=4`, `Dock=Left`, `Highlight`) + `lblCardTitulo*` (10pt semibold).
4. `tlpMain`: `Padding=12`; filas AutoSize/AutoSize/AutoSize/100 % (sin cambios de estructura).
5. Período: `tlpPeriodo` 3 cols con labels arriba (mismos combos + mismos handlers `Periodo_SelectedIndexChanged`).
6. Rutas: mismos `txt*` + `btnSeleccionar*` con `Text="..."`, `Size≈32×28`, `Anchor=Right`; `txt*` `Anchor=Left|Right`; tips "Examinar…" en `toolTipRutas`.
7. Ejecución: misma fila ASE; botones con `FlatStyle`/`BackColor`/`ForeColor`/`FlatAppearance` según §2.5 (textos intactos); progreso igual HU-18.
8. Resultado: `lblEstadoHumano` + `lblResumenUnaLinea` (empty states §2.4) + `chkVerDetalle` (`Checked=false`, `Text="Ver detalle técnico"`) + `pnlDetalleTecnico` (`Visible=false`, `Dock=Fill`) conteniendo `txtLog` + `lblLineasLog` movidos con sus props intactas.
9. `TabIndex` según §2.2; `txtLog.TabStop=false`.
10. **Verificación U1:** build 0 warnings (el Designer roto no compila — el compilador es el primer test) + abrir la app: header slim con línea de acento, 4 cards blancas, ventana 960×680, redimensión sin solapes, Resultado colapsado por defecto. Funcionalidad aún HU-18 salvo que `chkVerDetalle` no hace nada hasta U2 (consignarlo) — el resumen humano mostrará empty states fijos.

### U2 — Code-behind only: collapse + resumen humano + extensión de habilitación (AC-PER-02, AC-RUT-02, AC-EJE-02/03, AC-RES-02/03/04)

Archivo: `Remuneracion.WinForms/Form1.cs` (ÚNICO; prohibido tocar el Designer en U2 salvo `TabIndex` de emergencia documentada).

1. `chkVerDetalle_CheckedChanged`: `pnlDetalleTecnico.Visible = chkVerDetalle.Checked; tlpMain.PerformLayout();` Trazabilidad: AC-RES-02.
2. Helper `ActualizarResumenHumano(string? lineaResumen = null)`: `lblEstadoHumano.Text = toolStripStatusLabel.Text;` + `lblResumenUnaLinea.Text = lineaResumen ?? <empty/último>`; invocarlo en cada sitio que hoy escribe `toolStripStatusLabel.Text` (idle/limpiar, "Procesando…", completado 1-ASE/5-ASE con formatos §2.4 usando variables YA existentes, cancelado, error plantilla, error catch, "Log copiado…"). Trazabilidad: AC-RES-03, AC-SLOP-05.
3. `SetControlesHabilitados`: agregar `chkVerDetalle.Enabled = habilitados;` (única extensión de firma/lógica permitida). Trazabilidad: AC-EJE-03.
4. `btnLimpiar_Click`: +2 líneas (`chkVerDetalle.Checked=false; ActualizarResumenHumano();` con empty states). Rutas/período intactos (G4 HU-18 vigente). Trazabilidad: AC-RES-04.
5. `AppendLogLine`, `Ruta_*`, `Periodo_*`, `ConfigurarSerilog`, `RunId`, `MostrarErrorUx`, `UltimoCodigoSalida`: INTACTOS (solo se les agrega la llamada espejo del paso 2 donde corresponda).
6. **Verificación U2:** build 0 warnings + smoke funcional: arranque colapsado (AC-RES-01); expandir/colapsar conserva contenido; Limpiar restaura colapso; Copiar pega idéntico en Notepad; ejecución corta con % y ASE vivos; error provocado (salida==plantilla) muestra guía HU-14 + status/resumen humanos + código solo en file-log; redimensionar 960→940→maximizado sin solapes.

### U3 — Polish anti-slop + verificación final + cierre (AC-SLOP-01..05, AC-REG-01, AC-BLD-01, AC-GIT-01)

1. Checklist premium (todo SÍ): ☐ 0 `GroupBox`; ☐ fondo window ≠ card (contraste visible); ☐ Ejecutar primario vs 3 secundarios (jerarquía a ojo + props §2.5); ☐ espaciado múltiplo de 4; ☐ 0 colores hardcodeados; ☐ 0 paths truncados sin remedio; ☐ 0 códigos técnicos visibles; ☐ empty states §2.4; ☐ foco visible + TabOrder §2.2; ☐ español administrativo.
2. Greps de cierre: `GroupBox` = 0; `FlatStyle` ≥ 4 matches (1 primario + 3 secundarios); `FromArgb|FromName` = 0; `(salida ` en `Form1.cs` = 0.
3. Regresión bloqueante: build 0/0 + Q1/Q2 5-ASE ±0.5 vs goldens/baseline (SHAs anotados) + `git status` WinForms-only (AC-GIT-01).
4. Smoke de redimensión: 960×680 → 940×660 (mínimo) → maximizado.
5. **Verificación U3 (cierre):** tabla AC-PER/AC-RUT/AC-EJE/AC-RES/AC-SLOP/AC-REG/AC-BLD/AC-GIT pasa-falla + SHAs Q1/Q2 + build 0/0. Reportar al Inge para aprobación. Sin commit.

---

## 5. Architecture Validation Certificate

| Principio | Veredicto | Justificación |
|---|---|---|
| SRP | ✅ | `chkVerDetalle_CheckedChanged` solo alterna visibilidad; `ActualizarResumenHumano` solo escribe 2 labels; cálculo/logging intactos en sus clases. |
| OCP | ✅ | Se extiende `SetControlesHabilitados` y `InitializeComponent`; `Procesador*`, `Locator`, catálogo, CLI: cero modificaciones. |
| DIP | ✅ | `Form1` sigue dependiendo de `IProcesadorRemuneracion/IProcesadorPeriodo` ctor-inyectados; lo nuevo (`Panel`, `CheckBox`, `FlatAppearance`) es plataforma WinForms, no dominio. |
| Best practices WinForms | ✅ | `Suspend/ResumeLayout`, `Anchor/Dock`, `SystemColors` (tema/alto contraste), `AutoEllipsis` en labels de resumen, `TabStop=false` en log, `PerformLayout` tras colapso, `BeginUpdate/EndUpdate` existentes intactos. |
| Performance | ✅ | Colapso = `Visible` O(1); espejo de status = 1 asignación de string por evento; sin timers ni layouts anidados nuevos salvo `tlpPeriodo` (1 nivel); ejecución sigue en `Task.Run`. |
| Seguridad | ✅ | Sin nuevos IO/handles: `Process`/`Clipboard` ya existen HU-18 con guards; el colapso no expone datos nuevos (el log ya era visible; ahora está un clic más lejos, no más expuesto). |

---

## 6. Riesgos y mitigaciones

| Riesgo | Prob. | Mitigación |
|---|---|---|
| Renombre `grp*`→`pnlCard*` rompe algo no visto | Muy baja | `Form1.cs` no referencia `grp*` (verificado V-lectura completa); `grep "grp[A-Z]"` post-U1 = 0 confirma |
| `BorderStyle.FixedSingle` en Panel se ve distinto al mock mental (borde SO) | Media | Aceptado: es el borde fino nativo (anti-slop); si en smoke se ve grueso, U3 acota a `Padding`+fondo sin cambiar tipos |
| Resumen 1-línea queda largo en 940 px | Baja | `AutoEllipsis` + ToolTip implícito del SO; el formato §2.4 es corto por diseño (sin paths: solo nombre de archivo) |
| Status espejo diverge del `StatusStrip` | Muy baja | Espejo literal en un helper único (`lblEstadoHumano.Text = toolStripStatusLabel.Text`); `grep toolStripStatusLabel.Text =` cuenta los call-sites (7) y cada uno llama al helper |
| Regresión Q1/Q2 difiere | Muy baja | Candado AC-GIT-01 + AC-REG-01; rollback = `git checkout -- Remuneracion.WinForms/` |
| El aire sobrante de la card colapsada molesta al Inge | Media | G1 lo declara aceptado; alternativa (fila AutoSize) documentada como follow-up fuera del plan |

---

*Plan generado por sdd-planner — Fuente Única de Verdad para HU-19. Implementar solo con aprobación del Inge. Sin código implementado en este documento.*
