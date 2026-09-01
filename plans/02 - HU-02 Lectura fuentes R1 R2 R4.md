# Plan 02 — HU-02: Lectura de las fuentes R1/R2/R4 (readers)

> **Historia:** HU-02 — Lectura de las 8 fuentes (readers) · alcance reducido a R1/R2/R4 (obligatorios de Fase 1)
> **Fuente única de verdad:** `requirements/Fase1-Requerimientos.md` (HU-02, §4, §6 reglas 1–4) · `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` §5.1, §6, §9 (iteraciones 1.3, 1.4, 1.5)
> **Estado:** PLAN — pendiente de aprobación para implementación
> **Fecha:** 2026-09-01
> **Alcance:** Implementar `ExcelDataReaderRecaudoReader` (stub → real) para R1, R2, R4 usando ExcelDataReader 3.9.0. **NO** incluye HU-03 (cálculo) ni HU-04 (escritura).

---

## 0. Clarification Gate

| Pregunta / Ambigüedad | Resolución | Fuente de verificación |
|---|---|---|
| ¿Posiciones de fila/columna fijas? | **No.** Parsing 100 % dinámico por contenido de celdas (regla de negocio #11). | Archivos reales en `Docs/Insumos/REMUNERACION 2026071/1-Promoambiental/` — verificados en este plan |
| ¿`DetallePorComponente` (R1) se puebla en esta HU? | **No.** Fuera de alcance: lo consume la hoja `Reporte Componentes R1` (HU-03/HU-04). Este plan llena `TotalOportuno`, `Extemporaneo`, `NombreAse`. | Requisitos §6 regla 8; alcance HU-02 §5 |
| ¿De dónde sale `NombreAse` si la firma es solo `rutaArchivo`? | Se deriva del **nombre de la carpeta padre** del archivo (ej. `1-Promoambiental`). Fallback: nombre base del archivo. Decisión de diseño D3. | Estructura real de carpetas §5.2 Propuesta |
| ¿`Infrastructure` dependen de Serilog para logging? | **No.** Decisión de diseño D2. El logging queda en el llamador (WinForms). | Stack del proyecto: Serilog solo en `Remuneracion.WinForms` |
| ¿Hay test project? | No existe aún. Se valida con un **harness de verificación por consola** temporal (decisión D5); xUnit + golden test llegan en HU-05. | Project context «No hay test project aún» |

Sin ambigüedades bloqueantes: plan generado directamente.

---

## 1. PROPOSE — Propuesta

### 1.1 Por qué este cambio

Los readers de Excel son **stubs que lanzan `NotImplementedException`** (`Remuneracion.Infrastructure/Excel/ExcelDataReaderRecaudoReader.cs`). HU-03 (cálculo del CONSOLIDADO) y HU-04 (escritura de plantilla) dependen de los datos que estos readers deben producir. Sin esta implementación, la cadena completa de Fase 1 no puede ejecutarse.

Esta historia entrega la lectura de las **3 fuentes obligatorias** de la Fase 1 (R1, R2, R4) contra archivos fuente **reales** ya presentes en `Docs/Insumos/REMUNERACION 2026071/` (5 ASE), con valores de referencia verificados (Sección 4).

### 1.2 Trazabilidad

| Fuente | Referencia | Contenido |
|---|---|---|
| `requirements/Fase1-Requerimientos.md` | **HU-02** (§5) | Readers de las 8 fuentes; criterios de aceptación R1/R2/R4 |
| `requirements/Fase1-Requerimientos.md` | §6 reglas **1, 2, 3, 4, 11** | TOT_OPT (Componente+Total→F), R2 (GrandTotal−SERV_ESP_K), EXTEMP (Mes→F), R4 (última/Total→col 4 negativo), parsing dinámico |
| `requirements/Fase1-Requerimientos.md` | §4 mapa insumos | R1/R2/R4 → HU-02 y HU-04 |
| Propuesta | **§5.1** | Nombres exactos de archivos y qué extrae cada uno |
| Propuesta | **§6.1 / §6.2** | Principios (lectura directa de fuente, separación de capas); `RecaudoReader: lee R1/R2/R4` en Infrastructure |
| Propuesta | **§9 iteraciones 1.3, 1.4, 1.5** | Lector R1 (TOT_OPT/EXTEMP), R2 (Grand Total/SERV_ESP_K), R4 (total reversiones) |

### 1.3 Alcance

**In (scope):**
- Implementar `LeerR1`, `LeerR2`, `LeerR4` en `ExcelDataReaderRecaudoReader`.
- Helpers privados de lectura (apertura, barrido de filas, normalización de texto, parsing numérico).
- Manejo de errores con excepciones de dominio existentes.
- Harness de verificación por consola (validación contra archivos reales).

**Out of scope (explícitamente NO):**
- Readers de banco, reversados, balance, saldos por nota, retribución negativa (Q2/Fase 2).
- `DetallePorComponente` de R1 (lo consume HU-03/HU-04).
- Motor de cálculo (`ICalculoRemuneracion`) — HU-03.
- Escritura de plantilla (`IPlantillaWriter`) — HU-04.
- Proyecto de tests xUnit — HU-05.
- Cambios a contratos `IRecaudoReader` ni a modelos Core.

### 1.4 Criterios de éxito

- [ ] `dir` build limpio con **0 warnings** (`dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx`).
- [ ] El harness de verificación reporta PASS para los 5 valores de referencia (Sección 4) con tolerancia **±0,5**.
- [ ] Sin archivo → `ArchivoFuenteNoEncontradoException`; estructura inesperada → `CalculoInvalidoException`.
- [ ] No se leen fórmulas ni se escribe nada (solo lectura).

### 1.5 Riesgos

| Riesgo | Prob. | Mitigación |
|---|---|---|
| Variación de estructura entre quincenas/ASE (filas/columnas dinámicas) | Media | Búsqueda por contenido y encabezados; nunca por índice fijo de fila |
| ASE4 sin columna "Especiales" | Media | Chequeo dinámico de encabezado `Especiales`; ausente → 0 |
| Tipos numéricos con decimales (Excel→double) | Baja | `Convert.ToDecimal` con `InvariantCulture` |
| Archivos fuente con tilde/encoding en nombre (`ReversiónPorComponente`) | Baja | Ya resuelto por `ArchivoFuenteLocator` (case-insensitive); el reader recibe rutas resueltas |

---

## 2. DESIGN — Diseño técnico

### 2.1 Decisiones de diseño (registradas)

| # | Decisión | Justificación | Consecuencia |
|---|---|---|---|
| **D1** | Clase de lectura única (`ExcelDataReaderRecaudoReader`) con helpers privados `static` | No crear una clase helper nueva en Infrastructure: el único consumidor es este reader; evita sobre-ingeniería en Fase 1 | Mantener cohesionado el archivo; sin cambios de superficie pública |
| **D2** | **Sin logging en Infrastructure** (no depender de Serilog) | Serilog vive solo en `Remuneracion.WinForms`; añadirlo rompería la separación de capas | Los readers son mudos; el llamador (Form1/HU-05) loguea entradas/salidas |
| **D3** | `NombreAse` se deriva del nombre de la carpeta padre (patrón `N-Nombre`) | La firma del contrato solo recibe la ruta; la carpeta es fuente confiable (`1-Promoambiental`) | `Path.GetFileName(Path.GetDirectoryName(ruta))`; fallback nombre de archivo |
| **D4** | Barrido completo a memoria: `List<object?[]>` de filas, luego consulta LINQ | Archivos pequeños (<40 filas); simplifica búsquedas por contenido y reutiliza la misma fila para múltiples lookups (R2) | Una sola pasada de lectura; sin re-lecturas |
| **D5** | Verificación con **harness de consola temporal** (`Herramientas/VerificadorRecaudo`) en lugar de tests xUnit | No hay test project; HU-05 introduce xUnit + golden. El harness da evidencia fresca ejecutable hoy y será descartado | Proyecto `net10.0` de consola, **sin paquetes**, NO se agrega al `.slnx`; se elimina al llegar HU-05 |
| **D6** | `ServEspK` se busca por **encabezado** (`Especiales`, case-insensitive, trimmed), no por columna fija 11 | Regla de negocio #2; ASE4 tiene otra columna 11 | Verificado: Promoambiental 202607-1 SÍ tiene col K `Especiales` (57261.70) — la detección dinámica la encontró; ASE4 puede variar el encabezado y debe dar 0 |

### 2.2 Enfoque de implementación (contract-first, sin rediseñar)

Mantener la firma `IRecaudoReader` intacta. Implementar en `ExcelDataReaderRecaudoReader`:

#### 2.2.1 Helpers compartidos

```csharp
// 1) Apertura: valida existencia → ArchivoFuenteNoEncontradoException;
//    abre con FileShare.ReadWrite (permite archivo abierto en Excel); crea reader.
private static IExcelDataReader AbrirReader(string rutaArchivo)

// 2) Barrido: lee la PRIMERA hoja (todas tienen una sola: "Sheet1").
//    Devuelve List<object?[]> — filas → celdas (normalizadas de tamaño = FieldCount).
private static List<object?[]> LeerFilas(IExcelDataReader reader)

// 3) Texto normalizado: null → ""; string → Trim(); otro tipo → Convert.ToString(InvariantCulture)
private static string CeldaTexto(object? valor)

// 4) Número culture-safe: null/"" → 0m; double/decimal → Convert.ToDecimal(InvariantCulture)
private static decimal CeldaNumero(object? valor)
```

> Nota API ExcelDataReader 3.9.0: `reader.GetValue(i)` devuelve `object?` (double para números, string para texto); `reader.Name` da el nombre de la hoja. Abrir SOLO la primera `Result`. No se requiere `CodePagesEncodingProvider` para `.xlsx` (defensivo, sin agregar paquete).

#### 2.2.2 `LeerR1` — Recaudoporcomponente

| Valor | Criterio de búsqueda (0-based) | Columna | Real 202607-1 (Promoambiental) |
|---|---|---|---|
| `TotalOportuno` | fila con `A=="Componente"` **y** `B=="Total"` (case-insensitive) | **F (index 5)** | `19556118465.99` |
| `Extemporaneo` | **primera** fila con `B=="Mes"` | **F (index 5)** | `19549786950.62` |

Reglas críticas:
- NO usar la fila `A=="Total" && B vacío` (suma de todo el reporte) ni la sección `Subsidio/Contribucion`.
- La fila `Componente/Total` aparece **después** de la fila `Mes` — buscar por contenido, no por orden.
- `NombreAse` = carpeta padre (D3). `DetallePorComponente` queda **vacío** (fuera de alcance).

#### 2.2.3 `LeerR2` — RerpoteDetalleSaldosaFavor

| Valor | Criterio de búsqueda (0-based) | Columna | Real 202607-1 (Promoambiental) |
|---|---|---|---|
| `GrandTotal` | fila con `A=="Total"` **y** `B` vacío — NO `Subs/Cont`/`Componente` con `B=="Total"` | **E (index 4)** | `54273647.38` |
| `TieneColumnaEspeciales` | ¿existe columna cuyo **encabezado** (fila de headers) normalizado == `"Especiales"`? | — | `true` (col K, index 10) |
| `ServEspK` | valor de esa columna en la **fila GrandTotal**; ausente → `0` | — | `57261.70` |

> ⚠️ **Corrección de verificación (2026-09-01):** el archivo real de 202607-1 (Promoambiental) **SÍ** contiene la columna `Especiales` en K (index 10) con valor `57261.70` en la fila GrandTotal. El análisis inicial que registraba `false/0` leyó solo hasta col J y estaba equivocado. La decisión D6 (detección dinámica por encabezado) se valida tal cual: el código la detectó correctamente. `TotalOportuno = 54273647.38 − 57261.70 = 54216385.68`.

- **Localizar fila de encabezados**: primera fila que contiene `"Total"` **y** `"Componente TDF"` en cualquier celda (en reales: fila index 3).
- **Búsqueda de `Especiales`**: iterar esa fila de encabezados comparando `CeldaTexto` con `"especiales"` (**OrdinalIgnoreCase**) → índice de columna; si no hay match → `TieneColumnaEspeciales=false`, `ServEspK=0`.
- **`TotalOportuno` es computado** (`GrandTotal - ServEspK`) en el modelo — NO asignarlo manualmente.
- `NombreAse` = carpeta padre (D3).

⚠️ Ambiguos reales relevantes: existen filas `[Subs/Cont, Total, ...]` y `[Componente, Total, ...]` — el criterio `A=="Total" && B vacío` las excluye.

#### 2.2.4 `LeerR4` — ReversiónPorComponente

| Valor | Criterio de búsqueda (0-based) | Columna | Real 202607-1 (Promoambiental) |
|---|---|---|---|
| `TotalReversiones` | fila con `A=="Total"` **y** `B` vacío (fila tope; la de `OCCIDENTE/Total` coincide pero no es la única) | **D (index 3)** | `-12054255.65` |

- **NO volver a negar**: el valor ya viene negativo en la fuente.
- `NombreAse` = carpeta padre (D3).

### 2.3 Manejo de errores

| Condición | Excepción | Mensaje debe incluir |
|---|---|---|
| `rutaArchivo` null | `ArgumentNullException` (patrón existente del stub) | — |
| Archivo no existe / no accesible | `ArchivoFuenteNoEncontradoException` (Core) | Ruta completa + reporte esperado |
| Fila buscada no encontrada (TOT_OPT, EXTEMP, GrandTotal, Total R4) | `CalculoInvalidoException` (Core) | Reporte + criterio buscado + filas leídas (resumen) |
| Celda de valor no numérica (esperada numérica) | `CalculoInvalidoException` (Core) | Fila/columna + contenido real |
| Archivo sin hojas o vacío | `CalculoInvalidoException` (Core) | Reporte + «sin datos» |

Todas las excepciones de dominio **ya existen** en `Remuneracion.Core/Exceptions/` — no crear nuevas.

### 2.4 Arquitectura resultante

```
Remuneracion.WinForms (log/Serilog — D2)
        │ llama
        ▼
Remuneracion.Infrastructure/Excel/ExcelDataReaderRecaudoReader  ← ESTE PLAN
        │ usa (ExcelDataReader 3.9.0, ya referenciada)
        ▼
Remuneracion.Core (Models + Exceptions + IRecaudoReader)  ← sin cambios en este plan
```

Dependencias: **ninguna** nueva (no se agregan paquetes; `System.Data` viene con ExcelDataReader).

### 2.5 Visual Design Intent

No aplica: HU-02 no contiene tareas de UI. La UI existente (`Form1`) no se modifica en este plan.

---

## 3. Estrategia de pruebas

> No hay test project en la solución (llega en HU-05). Se valida con harness de consola temporal + valores de referencia reales.

### 3.1 Valores de referencia (verificados contra archivos físicos — 1-Promoambiental, período 202607-1)

| # | Fuente | Archivo real | Valor esperado | Dónde se lee |
|---|---|---|---|---|
| T1 | R1 | `Recaudoporcomponente_...20267161653925.xlsx` | **19556118465.99** | TOT_OPT — A=`Componente`, B=`Total`, col F |
| T2 | R1 | ídem | **19549786950.62** | EXTEMP — B=`Mes`, col F |
| T3 | R2 | `RerpoteDetalleSaldosaFavor_...202671616325453.xlsx` | **54273647.38** | GrandTotal — A=`Total` (B vacío), col E |
| T4 | R2 | ídem | **57261.70** (corregido vs análisis inicial) | ServEspK — col K `Especiales` en fila GrandTotal; TotalOportuno = 54273647.38 − 57261.70 = **54216385.68** |
| T5 | R4 | `ReversiónPorComponente_...2026716163221489.xlsx` | **-12054255.65** | A=`Total` (B vacío), col D — ya negativo |

Tolerancia: **±0,5** (regla de negocio #9).

### 3.2 Harness de verificación (`Herramientas/VerificadorRecaudo`)

- Proyecto de consola `net10.0`, **sin paquetes**, `ProjectReference` a Core + Infrastructure. **No se agrega al `.slnx`** (se descarta en HU-05).
- Ejecuta `LeerR1/LeerR2/LeerR4` contra los 3 archivos reales (rutas relativas a `Docs/Insumos/REMUNERACION 2026071/1-Promoambiental/`).
- Evalúa cada valor con `Math.Abs(esperado - real) <= 0.5m` y reporta `PASS/FAIL` por caso + resumen.
- Casos negativos: ruta inexistente → espera `ArchivoFuenteNoEncontradoException`; archivo real pero con fila buscada eliminada en copia temporal → espera `CalculoInvalidoException`.

```bash
dotnet run --project Herramientas/VerificadorRecaudo/Herramientas.VerificadorRecaudo.csproj
```

### 3.3 Criterios de aceptación HU-02 (requisitos §5) verificables por el harness

| Criterio del requisito | Verificación |
|---|---|
| R1 extrae TOT_OPT y EXTEMP «coincidiendo con fila col1=Componente/col2=Total y sección Extemporáneo, ±0,5» | Casos T1, T2 |
| R2 Grand Total y SERV_ESP_K «según encabezado real de col. 11 (ASE4 → 0)» | Casos T3, T4 (+ prueba dinámica: copia con encabezado `Especiales` insertado → debe leerlo) |
| R4 «última fila, columna 4 (valor negativo)» | Caso T5 |

---

## 4. SPEC — Criterios de aceptación y trazabilidad

### 4.1 Escenarios de aceptación (HU-02 delimitado a R1/R2/R4)

| ID | Escenario (dado/cuando/entonces) | Traza a |
|---|---|---|
| AC-1 | Dado archivo R1 válido, cuando `LeerR1`, entonces `TotalOportuno` = fila A=`Componente`∧B=`Total` col F y `Extemporaneo` = 1.ª fila B=`Mes` col F, ambos dentro de ±0,5 vs manual | HU-02 CA-1; Reglas 1, 3; Propuesta 1.3 |
| AC-2 | Dado archivo R2 válido, cuando `LeerR2`, entonces `GrandTotal` = fila A=`Total`(B vacío) col E, `TieneColumnaEspeciales`/`ServEspK` según encabezado real, `TotalOportuno = GrandTotal − ServEspK` | HU-02 CA-2; Regla 2; Propuesta 1.4 |
| AC-3 | Dado archivo R4 válido, cuando `LeerR4`, entonces `TotalReversiones` = fila A=`Total` col D **sin re-negar** | HU-02 CA-3; Regla 4; Propuesta 1.5 |
| AC-4 | Dado archivo inexistente, cuando cualquier `LeerR*`, entonces `ArchivoFuenteNoEncontradoException` con ruta en mensaje | Requisitos §8 R5 |
| AC-5 | Dado archivo con estructura inesperada (fila clave ausente), cuando leer, entonces `CalculoInvalidoException` descriptiva | Requisitos §8 R5; Regla 11 |
| AC-6 | Dado archivo cuyo encabezado de col. 11 es `Componente TCS` (ASE4), cuando `LeerR2`, entonces `ServEspK = 0` | HU-02 CA-2; Regla 2 |
| AC-7 | La solución compila con **0 warnings** | HU-01/CA general; skill verification-before-completion |
| AC-8 | El reader no escribe archivos ni depende de Serilog | Principios §6.1; decisión D2 |

### 4.2 Matriz de trazabilidad completa

| Entregable | Requisito | Propuesta | Evidencia de verificación |
|---|---|---|---|
| Reader R1 (TOT_OPT/EXTEMP) | HU-02, reglas 1, 3, 11 | §5.1 R1, §9 1.3, §6.2 RecaudoReader | Casos T1/T2 harness |
| Reader R2 (GrandTotal/ServEspK) | HU-02, regla 2, 11 | §5.1 R2, §9 1.4 | Casos T3/T4 harness + AC-6 |
| Reader R4 (TotalReversiones) | HU-02, regla 4, 11 | §5.1 R4, §9 1.5 | Caso T5 harness |
| Manejo de errores | §8 R5 | §6.1 trazabilidad | AC-4, AC-5 |

---

## 5. TASKS — Desglose de implementación

| # | Tarea | Depende de | Archivos | Criterio de salida |
|---|---|---|---|---|
| T1 | Helpers compartidos: `AbrirReader` (con `ArchivoFuenteNoEncontradoException`), `LeerFilas` (1.ª hoja), `CeldaTexto`, `CeldaNumero` | — | `ExcelDataReaderRecaudoReader.cs` | Compila; helpers cubren null/vacío/InvariantCulture |
| T2 | Implementar `LeerR1` (TOT_OPT + EXTEMP + NombreAse por carpeta padre) | T1 | ídem | Casos T1/T2 del harness PASS |
| T3 | Implementar `LeerR2` (GrandTotal + detección dinámica de columna `Especiales`) | T1 | ídem | Casos T3/T4 PASS + AC-6 (copia con encabezado insertado) |
| T4 | Implementar `LeerR4` (TotalReversiones col D sin re-negar) | T1 | ídem | Caso T5 PASS |
| T5 | Crear harness `Herramientas/VerificadorRecaudo` (consola, sin paquetes) con 5 casos positivos + 2 negativos | T2, T3, T4 | `Herramientas/VerificadorRecaudo/` (nuevo) | `dotnet run` reporta 7/7 PASS |
| T6 | Build de la solución con **0 warnings** + limpieza (remover bloqueo del harness si hace falta) | T1–T5 | `Remuneracion.WinForms/Remuneracion.WinForms.slnx` | `dotnet build` limpio |

**Secuencia de dependencias:** `T1 → (T2, T3, T4) → T5 → T6`.

### Reglas de implementación

1. **No rediseñar contratos ni modelos** — implementar contra `IRecaudoReader` y los modelos existentes tal cual.
2. **No agregar paquetes NuGet** — ExcelDataReader 3.9.0 ya referenciada.
3. **Parsing dinámico por contenido** — nunca índices de fila fijos (regla 11).
4. **Búsquedas case-insensitive + trim** sobre texto normalizado.
5. **0 warnings** después de cada tarea.
6. **No tocar UI (Form1) ni OpenXML writer** en este plan.
7. `NombreAse` derivado de carpeta padre (D3) — no inventar otros datos.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Cumplimiento |
|---|---|
| **S** — SRP | Cada método de lectura tiene una responsabilidad única; helpers separados por función (apertura, barrido, normalización, número) |
| **O** — OCP | Estructura variable absorbida por búsquedas por contenido/encabezados; nuevos patrones de archivo no requieren cambios en el flujo central |
| **L** — LSP | Sin herencia; el reader se consume vía interfaz `IRecaudoReader` — cualquier implementación es intercambiable |
| **I** — ISP | La interfaz existente es cohesiva (3 métodos de lectura); no se altera |
| **D** — DIP | `Infrastructure` implementa contratos de `Core`; `Core` no conoce nada de Excel |

### 6.2 Best Practices

- `using` + `IDisposable` para stream y reader (ExcelDataReader implementa `IDisposable`).
- `FileShare.ReadWrite` al abrir: tolera archivos fuente abiertos en Excel.
- `CultureInfo.InvariantCulture` en toda conversión numérica.
- Excepciones de dominio existentes, con mensajes descriptivos (sin excepciones genéricas).
- `ArgumentNullException.ThrowIfNull` coherente con el stub previo.
- Verificación con evidencia fresca (harness ejecutable), no afirmación sin ejecutar.

### 6.3 Performance

- **Una sola pasada** de lectura por archivo (barrido completo → consultas LINQ en memoria).
- Volúmenes mínimos (<40 filas × 10 columnas); sin streaming innecesario ni dobles lecturas.
- `LeaveOpen`/dispose correcto; sin fugas de archivos handles (bloqueo de archivo evitado con `FileShare.ReadWrite`).

**Veredicto de arquitectura:** APROBADO — el diseño es incremental, no introduce acoplamientos nuevos y respeta la separación de capas del proyecto.

---

## 7. Rollback

- El cambio es **aditivo y local** (archivo de reader + harness temporal). 
- Revertir = restaurar `ExcelDataReaderRecaudoReader.cs` al estado stub (git revert de los commits de HU-02). Ningún contrato ni modelo se modifica, por lo que no hay impacto colateral en Core/WinForms.
- El harness se elimina en HU-05 junto con la llegada del proyecto xUnit; mientras tanto vive fuera del `.slnx`, por lo que no afecta el build de la solución.

---

> **Control de cambios:** este plan cubre exclusivamente la implementación de lectura R1/R2/R4 (HU-02, alcance Fase 1). Readers Q2, cálculo, escritura y golden test corresponden a HU-03/HU-04/HU-05 y requieren planes propios.