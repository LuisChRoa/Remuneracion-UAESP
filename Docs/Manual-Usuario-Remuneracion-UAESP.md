# Manual de Usuario — Automatización de la Remuneración Quincenal UAESP

> **Norma de referencia:** Resolución UAESP 27 de 2018 — Reglamento Comercial y Financiero del servicio de aseo.
> **Documento rector:** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` (Fase 3, it. 3.4 — "Manual de usuario y technical").
> **Versión de la aplicación:** 1.0 (Fase 3 cerrada; HU-01..HU-17).

---

## 1. Qué hace la aplicación y qué NO hace

### 1.1 Qué hace

La aplicación automatiza la elaboración de la **remuneración quincenal** del servicio de aseo de Bogotá: distribuye el recaudo recibido y aplicado entre los cinco operadores (ASE) y lo asigna a cada componente tarifario, produciendo el archivo Excel consolidado `Remuneración AAAAMM-# Total.xlsx`.

Concretamente:

1. **Lee** los archivos fuente del período desde la carpeta de cada ASE (`Recaudoporcomponente`, `RerpoteDetalleSaldosaFavor`, `ReversiónPorComponente`, más los reportes de banco, balance y ajustes).
2. **Calcula** el CONSOLIDADO por ASE: TOT_OPT, R2 Total Oportuno, EXTEMP, Reversión R4 y AJUSTES-SF-T (solo 2.ª quincena).
3. **Escribe** los valores en las celdas hoja (leaf) de la plantilla copiada, **pegando solo valores** y sin tocar ninguna fórmula.
4. **Valida** la coherencia entre hojas con tolerancia **±0.5** por redondeo.
5. **Registra** cada paso de la ejecución en un log con un identificador de corrida (RunId).

### 1.2 Qué NO hace

- **No calcula ni escribe fórmulas**: las fórmulas del workbook las recalcula Excel al abrir el archivo (la app solo pega valores; OpenXML no recalcula).
- **No diligenciar la hoja INTERVENTORIA**: es un insumo externo anual que vive en la plantilla (ver §9).
- **No genera el flujo bancario** (`Proceso de Recaudo.docx`): ese proceso es previo y produce los archivos fuente.
- **No inserta ni elimina filas** en la plantilla; si la plantilla no alcanza, falla de forma explícita (fail-fast).
- **No es un instalador**: se publica como aplicación .NET framework-dependent (ver §2.2).

---

## 2. Requisitos y despliegue

### 2.1 Requisitos

| Requisito | Detalle |
|---|---|
| Sistema operativo | Windows (la UI es Windows Forms) |
| .NET | SDK **.NET 10** para compilar/ejecutar desde código (`dotnet --version`) o runtime .NET 10 para ejecutar una publicación |
| Excel | **Obligatorio para la Capa B** (recalcular y comparar el archivo de salida). La aplicación NO usa Excel/COM: solo genera el archivo. |

### 2.2 Despliegue (publicación framework-dependent)

La entrega se publica **sin instalador** (decisión HU-17 §2.6):

```bash
# UI
dotnet publish Remuneracion.WinForms/Remuneracion.WinForms.csproj -c Release

# CLI
dotnet publish Remuneracion.Cli/Remuneracion.Cli.csproj -c Release
```

Los ejecutables quedan en `bin/Release/net10.0-windows/publish/` (UI) y `bin/Release/net10.0/publish/` (CLI). Se distribuyen junto con sus DLL y un equipo con el **runtime .NET 10** instalado. Un instalador/setup queda explícitamente fuera; se re-evalúa solo si las pruebas de aceptación lo exigen.

### 2.3 Estructura de carpetas que la app espera

```
REMUNERACION AAAAMM{Q}/                  # Q = 1 o 2 (ej.: REMUNERACION 2026071)
├── 1-Promoambiental/                    # prefijos exactos de carpeta por ASE
├── 2-Lime/
├── 3-Ciudad Limpia/
├── 4-Bogotá Limpia/
└── 5-Área Limpia/
```

Dentro de cada carpeta de ASE, los archivos fuente se reconocen por **prefijo** (el resto del nombre, con fechas y timestamps, es variable):

| Reporte | Prefijo del archivo |
|---|---|
| R1 — Recaudo por componente | `Recaudoporcomponente_*.xlsx` |
| R2 — Detalle de saldos a favor | `RerpoteDetalleSaldosaFavor_*.xlsx` |
| R4 — Reversión por componente | `ReversiónPorComponente_*.xlsx` (fallback `ReversionPorComponente_*`) |
| Reporte por banco | `ReportePagosxBanco_*.xlsx` |
| Balance de subsidios/contribuciones | `R4-BalanceSubsidioyContribuciones_*.xlsx` (fallback variante `-Optimizado_`) |
| Saldos por nota (Q2) | `SaldosaFavorAplicadosPorNotas_*.xlsx` |
| Retribución negativa (Q2) | `RetribuciónNegativa_*.xlsx` |

**Plantilla de origen:** el archivo `Remuneración AAAAMM-# Total.xlsx` completado de referencia (el "canónico"). La app lo **copia** y escribe en la copia; la plantilla **jamás se modifica**.

---

## 3. Modo UI (Windows Forms)

### 3.1 Pantalla y campos

| Campo | Descripción |
|---|---|
| **Año / Mes / Quincena** | Período a liquidar. Quincena 1 = días 1–15; quincena 2 = días 16–31. |
| **Carpeta de fuentes** | Carpeta del período (`REMUNERACION AAAAMM{Q}`) con las 5 carpetas de ASE. |
| **Plantilla** | Archivo plantilla de origen (canónico). |
| **Carpeta de salida** | Carpeta donde se creará `Remuneración AAAAMM-# Total.xlsx` (se nombra automáticamente). |
| **Modo 5 ASE / ASE único** | Checkbox "5 ASE" activa los 5 operadores; desmarcado procesa el ASE seleccionado en el combo. |

### 3.2 Pasos

1. Seleccione año, mes y quincena del período.
2. Seleccione la carpeta de fuentes (`...` junto a "Carpeta de fuentes").
3. Seleccione la plantilla canónica (`...` junto a "Plantilla").
4. Seleccione la carpeta de salida (`...` junto a "Carpeta de salida").
5. Marque "5 ASE" o elija un ASE único en el combo.
6. Pulse **Ejecutar**.

### 3.3 Qué ocurre durante la ejecución

- El panel de log muestra los hitos por ASE (TOT_OPT, R2, EXTEMP, R4, resúmenes por empresa, banco, BCE, AJUSTES, DetRetri, INTERVENTORIA y VALIDACIONES).
- La barra de progreso avanza (máx. 42 hitos en modo 5 ASE, 8 en modo 1 ASE).
- Si el archivo de salida ya existe, la UI **pregunta** si se sobrescribe.
- Al terminar, la barra de estado muestra `Completado (salida 0)`. Si algo falla, muestra el código del catálogo y el código de salida.

### 3.4 Salida

El archivo resultante `Remuneración AAAAMM-# Total.xlsx` se genera en la carpeta de salida con la nomenclatura oficial. **Ábralo en Excel** para que recalcule las fórmulas (Capa B — ver `Docs/Instructivo-Capa-B.md`).

---

## 4. Modo CLI (línea de comandos)

### 4.1 Flags exactos

```
Uso: Remuneracion.Cli --periodo <AAAAMM|AAAAMMQ> --carpeta <dir-periodo> --plantilla <xlsx> --salida <dir>
                     [--ase N | --cinco-ase] [--quincena 1|2] [--sobrescribir] [--help]

  --periodo AAAAMM[Q]  Período (p. ej. 2026071). Si usa AAAAMM, --quincena es obligatoria.
  --quincena 1|2       Quincena (override/complemento del período).
  --carpeta <dir>      Carpeta del período que contiene las 5 carpetas de ASE.
  --plantilla <xlsx>   Plantilla de origen (se copia; jamás se modifica).
  --salida <dir>       Carpeta de salida (se crea si falta; el archivo se nombra por el período).
  --ase N              Procesa un solo ASE (1..5). Excluye --cinco-ase.
  --cinco-ase          Procesa los 5 ASE (predeterminado).
  --sobrescribir       Sobrescribe la salida si ya existe.
  --help, -h           Muestra esta ayuda y sale con 0.
```

Reglas de uso:
- `--periodo` es **obligatorio** y acepta `AAAAMM` (6 dígitos) o `AAAAMMQ` (7 dígitos con quincena 1 o 2). Con `AAAAMM`, `--quincena` es obligatoria.
- `--carpeta`, `--plantilla` y `--salida` son **obligatorios**.
- `--ase` y `--cinco-ase` son mutuamente excluyentes.
- El modo 1-ASE en quincena 2 falla por diseño con `ERR-VALIDACION` (salida 1): la ruta Q2 completa (AJUSTES-SF-T/DetRetri) vive en modo 5 ASE.

### 4.2 Ejemplos reales

```bash
# Q1 — 5 ASE (caso CT-Q1-5A)
Remuneracion.Cli --periodo 2026071 --carpeta "Docs/Insumos/REMUNERACION 2026071" --plantilla "Docs/Insumos/Remuneracion 202607-1 Total.xlsx" --salida "Docs/Insumos/Salidas" --cinco-ase --sobrescribir

# Q1 — ASE único 3 (caso CT-Q1-1A)
Remuneracion.Cli --periodo 2026071 --carpeta "Docs/Insumos/REMUNERACION 2026071" --plantilla "Docs/Insumos/Remuneracion 202607-1 Total.xlsx" --salida "Docs/Insumos/Salidas" --ase 3

# Q2 — 5 ASE (caso CT-Q2-5A; plantilla canónica Q2)
Remuneracion.Cli --periodo 2026072 --carpeta "Docs/Insumos/REMUNERACION 2026072" --plantilla "Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx" --salida "Docs/Insumos/Salidas" --cinco-ase --sobrescribir
```

Si ejecuta desde el código fuente:

```bash
dotnet run --project Remuneracion.Cli/Remuneracion.Cli.csproj -- --periodo 2026071 --carpeta "..." --plantilla "..." --salida "..." --cinco-ase
```

### 4.3 Línea de resultado (grepable)

Al terminar, el CLI imprime **una** línea final en stdout:

```
RESULTADO OK codigo=0 salida=<ruta-del-archivo> runId=<guid>
RESULTADO ERROR codigo=<ERR-…> salida=<N> runId=<guid>
```

Los hitos intermedios van a stdout; el detalle completo va al log con el mismo `runId`.

---

## 5. Códigos de salida

| Código | Significado | Causas típicas | Acción |
|---|---|---|---|
| **0** | OK | Proceso completado | Verificar Capa B |
| **1** | Validación | `ERR-VALIDACION`: una validación no cierra ±0.5; Q2 en modo 1-ASE (diseño) | Revisar el log con el RunId; usar modo 5 ASE en Q2 |
| **2** | Fuente o plantilla | `ERR-FUENTE-NO-ENCONTRADA` (falta carpeta/archivo), `ERR-PLANTILLA` (plantilla ausente o salida == plantilla), `ERR-FORMATO-FUENTE` (estructura inesperada) | Verificar carpeta/archivos/plantilla; usar salida distinta a la plantilla |
| **3** | Escritura | `ERR-ESCRITURA`: no se pudo escribir el workbook | Revisar permisos/bloqueo de archivo; reintentar (no queda parcial) |
| **4** | Inesperado / uso | Uso inválido (flag desconocido, período malformado, `--ase 6`, exclusión `--ase`+`--cinco-ase`, flag requerido ausente) o `ERR-INESPERADO` | Leer el mensaje de stderr + ayuda; corregir el comando |
| **5** | Cancelado | `WARN-CANCELADO`: la salida ya existe y no se usó `--sobrescribir` | Reejecutar con `--sobrescribir` o cambiar la carpeta de salida |

> En la UI, los códigos 0–5 se **registran** en la barra de estado y el log; la UI nunca termina el proceso con `Environment.Exit`.

---

## 6. Catálogo de errores y RunId

### 6.1 Catálogo (código → título → guía)

| Código | Título en la UI | Guía accionable |
|---|---|---|
| `ERR-FUENTE-NO-ENCONTRADA` | Archivo fuente faltante | Falta un archivo fuente del período: verifique la carpeta indicada, genere el reporte faltante y reintente. |
| `ERR-PLANTILLA` | Plantilla o ruta de salida no válida | Use una plantilla existente y una ruta de salida distinta a la plantilla. |
| `ERR-FORMATO-FUENTE` | Formato de archivo fuente no compatible | Revise el archivo, la hoja y la celda indicadas. |
| `ERR-VALIDACION` | La validación no cierra | La validación no cierra dentro de la tolerancia ±0.5. |
| `ERR-ESCRITURA` | Error al generar el archivo de salida | No quedó archivo parcial; reintente. |
| `WARN-CANCELADO` | Proceso cancelado | Proceso cancelado por decisión del usuario. |
| `ERR-INESPERADO` | Error inesperado | Búsquelo en el log con el RunId de esta ejecución. |

### 6.2 RunId — cómo filtrar una ejecución

Cada ejecución genera un **RunId** (GUID) que se adjunta a todos los eventos de esa corrida en el log. Para aislar una ejecución:

```powershell
Select-String -Path remuneracion_log_*.txt -Pattern "<runId>"
```

El mismo RunId aparece en la línea `RESULTADO OK/ERROR` del CLI, en la barra de estado de la UI y en cada evento del log.

---

## 7. El log

- **Archivo:** `remuneracion_log_.txt` con **rolling diario** (`remuneracion_log_20260909.txt`, …) y **retención de 30 días** (los archivos más antiguos se eliminan solos).
- **Nivel:** `Debug` (incluye hitos y detalle de validaciones).
- **Formato por evento:**

```
{Timestamp} [{Level:u3}] (RunId={RunId} Periodo={Periodo} AseId={AseId} Hoja={Hoja} Validacion={Validacion}) {Message}{Exception}{Properties}
```

- **Propiedades estructuradas** buscables: `RunId`, `Periodo`, `AseId`, `Hoja` (p. ej. `INTERVENTORIA`, `BCE SC POR FACT.`, `AJUSTES - SF-T`, `DetRetri2026072`, `REPORTE RECAUDO x BANCO`), `Validacion` (p. ej. `VALIDACION_ENEL`, `DetValiRetri`, `VALIDACION_TOTAL`), más cualquier propiedad adicional en `{Properties}`.
- La UI y el CLI usan la **misma configuración** (paridad garantizada por test).

---

## 8. Casos que NO son error

| Situación | Por qué es normal |
|---|---|
| `E59:E80` ≠ 0 en `REPORTE RECAUDO x BANCO` | Diferencias informativas entre anulado/reversado de la misma quincena; esperadas distintas de 0. |
| `D21` y `D9:D14` / `J9:J14` excluidos de las validaciones | Excluidos por decisión de diseño (HU-13); no son objetivos de validación. |
| Retribuciones vacías leídas como 0 | "Leído 0" es un cero legítimo de la fuente (no confundir con "slot ausente", que sí falla nombrando la celda). |
| Q2 en modo 1-ASE falla con salida 1 | Fail-fast por diseño: la ruta Q2 completa (AJUSTES-SF-T/DetRetri) vive en modo 5 ASE. |
| `AJUSTES-SF-T` en blanco en Q1 | Los ajustes solo aplican a la 2.ª quincena; en Q1 la app escribe 0. |
| El valor de una celda de visible cambia al abrir en Excel | Correcto: las fórmulas recalculan; la app solo pega valores (Capa B verifica el resultado post-Excel). |

---

## 9. INTERVENTORIA (insumo externo anual)

La hoja `INTERVENTORIA` (bloque K25:N32) es una **tabla anual estática** que **no se escribe ni se inventa**: es un insumo externo que la plantilla ya trae diligenciado y la aplicación la preserva intacta (hoja protegida).

| Rango | Contenido |
|---|---|
| K26:K30 | Valor oficial del mes por ASE |
| M26:M30 | Segunda quincena por ASE |
| N26:N30 | Primera quincena por ASE |
| K31 / M31 / N31 | Totales por columna (fórmula SUM) |
| K32 | Gran total = SUM(M31:N31) (fórmula) |

**Cómo diligenciarla (si un período nuevo la requiere):** edítela directamente en la plantilla **antes** de ejecutar la aplicación, con los valores oficiales del año. La aplicación verifica su estructura (assert estructural) y falla si la hoja cambia de forma inesperada. Si el valor cambia a mitad de año, el cambio es anual y no quincenal: se actualiza la tabla y se re-ejecuta.

---

## 10. Preguntas frecuentes (FAQ)

**¿La app reemplaza a Excel?** No. La app genera el archivo; Excel lo recalcula. La verificación final (Capa B) se hace en Excel contra el golden.

**¿Puedo ejecutar dos veces sobre la misma salida?** No sin confirmación: si la salida existe, la UI pregunta y el CLI exige `--sobrescribir` (sin él, salida 5).

**¿Qué pasa si falta un archivo fuente?** Fail-fast con `ERR-FUENTE-NO-ENCONTRADA` (salida 2) nombrando el ASE y el reporte. No se genera salida parcial certificada.

**¿Cómo sé que los valores escritos son los correctos?** Compare en Excel contra el golden del período (criterio ±0.5) siguiendo `Docs/Instructivo-Capa-B.md`.

**¿El `Informe AFaseo Recaudo` se tiene en cuenta?** No. Se ignora para el consolidado.

**¿Qué es el harness `Herramientas/VerificadorRecaudo`?** Una herramienta de verificación de los lectores (HU-02..HU-05, 24 casos) que corre por fuera de la solución; no es parte del flujo de producción.

---

## 11. Nota técnica (it. 3.4 "technical")

### 11.1 ADR: Serilog en Core

`Remuneracion.Core` referencia **Serilog** pero con alcance confinado: solo orquestadores (`Procesador*`) usan `LogContext`/logging; el modelo (Models/Constants/Interfaces/Exceptions) es puro, sin sinks ni configuración. Justificación HU-14: la trazabilidad por paso (CA-6) exigía logging en el orquestador y la CLI comparte la misma configuración. Si un día un consumidor de Core no puede llevar Serilog, se extrae la costura.

### 11.2 Geometría Q2 ≠ Q1

El template de quincena 2 tiene **layout distinto** al de quincena 1. El cell-map Q2 es hermano del mapa Q1, explícito por (`Ase.Id`, período). Visibles CONSOLIDADO:

| Bloque | Q1 (mapa HU-07) | Q2 (mapa HU-12) |
|---|---|---|
| TOT_OPT D9:D13 | R1!F46/… | R1!F53/F206/F343/F468/F558 |
| EXTEMP D47:D51 | R1!F48/… | R1!F55/F208/F345/F470/F560 |
| R2 D28:D32 | R2!E41/… | R2!E43/E139/E256/E358/E438 |
| R4 D66:D70 | R4!D67/… | R4!D73/D168/D205/D320/D355 |

Esto se descubrió en HU-12 (lección de método: ante un cambio de quincena, el primer T0 dumpea los visibles CONSOLIDADO antes de asumir continuidad de layout).

### 11.3 Hoja DetRetri/DetValiRetri

Las hojas `DetRetri2026071/2026072` y `DetValiRetri2026071/2026072` existen **por nombre** con sufijo de período (sheets 36/37 en ambos canónicos, no 93/94 como decía un borrador del Plan 12 — fe de erratas en `plans/12` §9). `DetRetri2026072!D9:D14` son **valores** editables (la app escribe enteros ROUND(D104:D108,0)); `DetRetri2026072!D23:D28` y `D32:D36` son **fórmulas** protegidas.

### 11.4 INTERVENTORIA D2b y L-Especiales D3a/D3b (HU-16)

- **D2b**: INTERVENTORIA es insumo externo declarado (sin fuente en insumos): hoja protegida intacta + assert estructural + log con `Hoja="INTERVENTORIA"`. Cero escrituras; nunca valor inventado.
- **D3a**: L-Especiales menores = celdas no-cero de la columna L de `Reporte Componentes R1` mapeadas por rol/ocurrencia y escritas en la misma pasada (113 celdas en Q1/Q2).
- **D3b**: celdas L sin fuente = estáticas, no se tocan.
- **M2**: ninguna fuente de saldos-nota trae columna Especiales → `TieneColumnaEspeciales = false` fijado por código; ausente = 0.

### 11.5 Veredicto D/E del BCE

En `BCE SC POR FACT.` (filas 3–7 por ASE): **D = CONTRIBUCION** (positiva), **E = SUBSIDIO** (negativa), **F = D+E** (TOTAL BSC), **H = SISTEMA** (redondeado) con diferencia I ≈ ±0.4. El veredicto T0 (D←Contribución F-fuente, E←Subsidio E-fuente) se fijó con el golden; `CONSOLIDADO J9:J13` y K/M calculan por fórmulas desde BCE F3:F7 (verificar post-Excel en Capa B).

### 11.6 Golden honesto (Capa A)

La Capa A compara **leafs escritos + visibles de dominio** contra el golden; nunca compara caché de fórmula de la salida contra el golden (OpenXML no recalcula; prohibido A5). Por eso la verificación de valores recalculados es **Capa B** (manual en Excel).

---

## 12. Integridad de insumos (SHAs de ancla)

Los cuatro archivos de referencia de HU-17 (goldens/canónicos) tienen SHA256 fijados como ancla de integridad. **Si un SHA difiere, deténgase y escale** (no "ajuste" el hash):

| Archivo | SHA256 |
|---|---|
| `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` (golden Q1) | `0B090E9C4851D86DAC534F2965E4A38393CC5BC09350C2A824FD48F7931C096F` |
| `Docs/Insumos/Remuneracion 202607-2 Total.xlsx` (golden Q2) | `584310105AC7223CE840833E3CB64E26F95B61907C9ECD959A0A632EEC21DCB1` |
| `Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx` (canónico Q2) | `95825422B32FBE9E2B60F35C9638E492455FB98CFA182976B657FAC3520E0B8C` |
| `Docs/Insumos/REMUNERACION 2026072/Plantilla  _ Remuneracion 202607-2 Total.xlsx` (control Q2, nunca oráculo) | `509BF210435138B3D36947582E295F27CE1F4D9452BEB1A8D752F3EF4F57529B` |
