# Plan 01 — Refinamiento y cierre de HU-01: Andamiaje, dominio y UI profesional

> **Historia:** HU-01 — Andamiaje, dominio y localizador de fuentes
> **Estado anterior:** Solución 3 capas compilable (0 warnings), dominio modelado, interfaces definidas, locator implementado, UI skeleton con stubs.
> **Objetivo:** Cerrar HU-01 con calidad senior: UI profesional con campos de período dinámicos, botones normalizados, layout responsivo, y code-behind limpio.
> **Fecha:** 2026-08-31

---

## 1. Problemas detectados en la UI actual (`Form1.Designer.cs`)

| # | Problema | Severidad | Ubicación |
|---|---------|-----------|-----------|
| P1 | **Periodo es un ComboBox único** con valores hardcodeados ("202605-1", etc.) — no es dinámico ni profesional | Alta | `cmbPeriodo` |
| P2 | **Botones con tamaños inconsistentes**: `btnSeleccionarCarpeta/Plantilla` = 110×23, `btnEjecutar` = 120×30 — ilegibles en algunos DPI | Alta | Botones |
| P3 | **Sin StatusStrip** — el estado se muestra en un Label suelto en Y=515, fuera de contexto | Media | `lblEstado` |
| P4 | **Sin layout responsivo** — al redimensionar la ventana, el log y los controles no se adaptan | Media | Form1 |
| P5 | **Espaciado inconsistente** — gaps de 30px entre filas, sin alineación a grilla | Media | Todo el form |
| P6 | **ComboBox ASE** no muestra el número de carpeta, solo el nombre — dificulta identificación | Baja | `cmbAse` |
| P7 | **Sin separación visual** entre secciones de entrada (parámetros) y área de resultado (log) | Baja | Form1 |

---

## 2. Diseño propuesto de la UI

### 2.1 Layout general

```
┌─ Remuneración Quincenal UAESP ─────────────────────────────────────┐
│                                                                     │
│  ┌─ Período ──────────────────────────────────────────────────────┐ │
│  │  Año: [ComboBox]  Mes: [ComboBox]  Quincena: [ComboBox]       │ │
│  └────────────────────────────────────────────────────────────────┘ │
│                                                                     │
│  ┌─ Archivos ─────────────────────────────────────────────────────┐ │
│  │  Carpeta fuentes: [TextBox .....................] [Examinar..] │ │
│  │  Plantilla:       [TextBox .....................] [Examinar..] │ │
│  └────────────────────────────────────────────────────────────────┘ │
│                                                                     │
│  ┌─ Ejecución ───────────────────────────────────────────────────┐ │
│  │  ASE: [ComboBox]    [▶ Ejecutar]                             │ │
│  │  [═══════════════════════════ progress bar ═══════════════]   │ │
│  └────────────────────────────────────────────────────────────────┘ │
│                                                                     │
│  ┌─ Registro de operación ────────────────────────────────────────┐ │
│  │  [TextBox multiline scrollable — log]                          │ │
│  │                                                                │ │
│  │                                                                │ │
│  └────────────────────────────────────────────────────────────────┘ │
│                                                                     │
│  StatusStrip: [Listo]                                    [v1.0.0] │
└─────────────────────────────────────────────────────────────────────┘
```

### 2.2 Controles y especificaciones

#### Sección Período

| Control | Tipo | Tamaño | Valores | Comportamiento |
|---------|------|--------|---------|----------------|
| `cmbAnio` | ComboBox (DropDownList) | 80×23 | 2024–2030 (generados dinámicamente) | Default = año actual |
| `cmbMes` | ComboBox (DropDownList) | 120×23 | Enero–Diciembre | Default = mes actual |
| `cmbQuincena` | ComboBox (DropDownList) | 100×23 | "1.ª Quincena", "2.ª Quincena" | Default = detectada por fecha |

**Lógica interna:** La app construye el código de período (`AAAAMM#`) a partir de los 3 campos. El usuario ve nombres legibles; el sistema maneja el formato interno.

#### Sección Archivos

| Control | Tipo | Tamaño | Comportamiento |
|---------|------|--------|----------------|
| `txtCarpetaFuentes` | TextBox (ReadOnly) | 420×23 | Ruta de la carpeta del período |
| `btnSeleccionarCarpeta` | Button | 110×28 | Abre `FolderBrowserDialog` |
| `txtPlantilla` | TextBox (ReadOnly) | 420×23 | Ruta del archivo plantilla |
| `btnSeleccionarPlantilla` | Button | 110×28 | Abre `OpenFileDialog` (.xlsx) |

#### Sección Ejecución

| Control | Tipo | Tamaño | Comportamiento |
|---------|------|--------|----------------|
| `cmbAse` | ComboBox (DropDownList) | 200×23 | "1 - Promoambiental", ..., "5 - Área Limpia" |
| `btnEjecutar` | Button | 130×28 | Inicia el procesamiento |
| `progressBar` | ProgressBar (Marquee) | Índice al进度 | Visible solo durante procesamiento |

#### Sección Log

| Control | Tipo | Tamaño | Comportamiento |
|---------|------|--------|----------------|
| `txtLog` | TextBox (Multiline, ReadOnly, ScrollBar=Both) | Ancho completo, altura fill | Dock=Fill, anchclado al StatusStrip |

#### StatusStrip

| Control | Tipo | Comportamiento |
|---------|------|----------------|
| `toolStripStatusLabel` | ToolStripStatusLabel | Muestra estado: "Listo", "Procesando...", "Completado" |
| `toolStripVersion` | ToolStripStatusLabel (Spring=false) | Muestra versión: "v1.0.0" |

### 2.3 Dimensiones del Form

| Propiedad | Valor |
|-----------|-------|
| `ClientSize` | 700×580 |
| `MinimumSize` | 650×500 |
| `StartPosition` | CenterScreen |
| `Text` | "Remuneración Quincenal UAESP — Fase 1" |

---

## 3. Cambios en `Form1.cs` (code-behind)

### 3.1 Inicialización

```csharp
// Generar años dinámicamente (5 atrás, 2 adelante desde DateTime.Now.Year)
// Llenar meses con nombres en español
// Detectar quincena actual por DateTime.Now.Day
// Llenar ASE desde CarpetasAse.Prefijos
```

### 3.2 Propiedad `PeriodoSeleccionado`

```csharp
// Construye el código interno AAAAMM# a partir de cmbAnio + cmbMes + cmbQuincena
// Ejemplo: 2026 + Julio + 1.ª → "2026071"
```

### 3.3.btnEjecutar_Click

- **Mantener como stub** por ahora (la lógica real viene en HU-02/03/04)
- Pero la UI debe responder: deshabilitar controles, mostrar progress, registrar en log, habilitar al terminar
- Mantener el mensaje informativo de que los readers son stubs

### 3.4 Normalización de botones

- Todos los botones de acción: **110×28** (altura mínima para legibilidad en 96 DPI)
- `btnEjecutar`: **130×28** (ligeramente más ancho por ser acción principal)
- `FlatStyle.Standard` (default) — no inventar estilos

---

## 4. Archivos a modificar

| Archivo | Cambio |
|---------|--------|
| `Remuneracion.WinForms/Form1.Designer.cs` | **Reescritura completa** del layout — nuevos controles, tamaños, posiciones |
| `Remuneracion.WinForms/Form1.cs` | Lógica de inicialización dinámica, `PeriodoSeleccionado`, stub de ejecución mejorado |

---

## 5. Criterios de aceptación del refinamiento

| # | Criterio | Verificación |
|---|----------|-------------|
| CA-R1 | Los 3 campos de período (Año/Mes/Quincena) funcionan independientemente | inspección visual |
| CA-R2 | El código de período se construye correctamente (ej: 2026 + Julio + 1 → "2026071") | inspección de código |
| CA-R3 | Todos los botones tienen altura consistente (28px) y texto legible | inspección visual |
| CA-R4 | El StatusStrip muestra el estado actual | inspección visual |
| CA-R5 | La ventana se redimensiona correctamente (log se expande) | prueba manual |
| CA-R6 | El form tiene `MinimumSize` para evitar layouts rotos | inspección de código |
| CA-R7 | La solución compila con **0 warnings** | `dotnet build` |
| CA-R8 | Los nombres de ASE muestran el número de carpeta ("1 - Promoambiental") | inspección visual |
| CA-R9 | Meses en español, años generados dinácticamente, quincena detectada por fecha | inspección visual |

---

## 6. Reglas de implementación

1. **No inventar estilos** — usar `FlatStyle.Standard`, colores del sistema, fuentes por defecto
2. **No agregar paquetes NuGet** — solo usar controles nativos de WinForms
3. **Maintener 0 warnings** — construir después de cada cambio significativo
4. **El botón "Seleccionar..." no cambia de nombre** — es el patrón estándar de Windows
5. **El ASE combo se llena desde `CarpetasAse.Prefijos`** — no hardcodear
6. **El log se mantiene como TextBox** — no es MVP funcional aún, solo visual

---

## 7. Tareas de implementación

| # | Tarea | Dependencia | Archivos |
|---|-------|-------------|----------|
| T1 | Reescribir `Form1.Designer.cs` con el nuevo layout (3 secciones + StatusStrip) | — | `Form1.Designer.cs` |
| T2 | Implementar inicialización dinámica en `Form1.cs` (años, meses, quincena, ASE) | T1 | `Form1.cs` |
| T3 | Implementar `PeriodoSeleccionado` y wire-up de los 3 ComboBox | T2 | `Form1.cs` |
| T4 | Implementar stub de ejecución con feedback visual (deshabilitar/habilitar, log, progress) | T2 | `Form1.cs` |
| T5 | Build verification — 0 warnings | T1-T4 | — |

---

> **Control de cambios:** Este plan cubre exclusivamente el refinamiento visual y de UX de la HU-01. La lógica de lectura/escritura/cálculo NO se toca aquí.
