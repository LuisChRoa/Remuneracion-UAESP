# Plan 07 — HU-07: Multi-ASE Fase 2 (5 ASE) sobre la ruta leaf certificada

> **Historia:** escalar el prototipo Fase 1 (1 ASE) al procesamiento de los 5 ASE del período, en UNA sola llamada de escritura, sin reabrir semántica HU-04/HU-05/HU-06.
> **Rector (irrenunciable):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §4 Fase 2 it. **2.1 Procesamiento multi-ASE (iterar las 5 carpetas ASE)**, §10 CA + tolerancia ±0.5, §6 arquitectura (WinForms → Infrastructure → Core, preservación de fórmulas, Serilog).
> **Alcance rector §9 Fase 2:** entra **solo 2.1**. Salen 2.2 conciliación por empresa, 2.3 REPORTE RECAUDO x BANCO, 2.4 BCE, 2.5 AJUSTES-SF-T, 2.6 DetRetri, 2.7 validaciones cruzadas — son HUs posteriores.
> **Continuidad:** `plans/05` y `plans/06` (ruta leaf, golden Capa A, `IProcesadorRemuneracion` single-ASE). HU-01..HU-06 cerradas.
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-08

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. Las decisiones están resueltas abajo; el apply espera aprobación explícita.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación (dump OpenXML real de `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` + lecturas de valores):**

| Hecho | Método | Consecuencia |
|---|---|---|
| `CONSOLIDADO_TOTAL RECAUDO` tiene **filas por ASE**: D9:D13 TOT_OPT, D28:D32 R2, D47:D51 EXTEMP, D66:D70 R4, D85:D89 AJUSTES (=0 en Q1), D104:D108 totales, D109=SUM(D104:D108) | Lectura A1:D115 + dump `<f>` de sheet18.xml | El consolidado ya es multi-ASE por diseño de plantilla |
| Cada fila CONSOLIDADO apunta a **filas distintas** de R1/R2/R4 (ver §2.1, tabla completa) | Dump de fórmulas: D10→R1!F176, D29→R2!E135, D67→R4!D161, etc. | Escribir 5 ASE = escribir 5 bloques, no replicar 1 bloque |
| Las hojas R1/R2/R4 **apilan los 5 bloques ASE verticalmente** en una sola hoja (headers B: PROMOAMBIENTAL r4, LIME r~79, CIUDAD LIMPIA r~210, BOGOTA LIMPIA r~350, AREA LIMPIA r~460 en R1; mismo patrón en R2/R4) | Lecturas columna B + valores por bloque | NO hay una hoja por ASE. Estrategia "una plantilla por ASE" queda descartada |
| R2 y R4 tienen **fórmulas visibles uniformes** por bloque (`E_vis=E_a+E_b-K_a`; `D_vis=D_x-P_x`) | Dump `<f>`: E41/E135/E247/E343/E413 y D67/D161/D198/D312/D347 | R2/R4 admiten mapa paramétrico por bloque |
| **R1 NO es uniforme**: ASE1 `F46=F25+F41-L25` (3 términos); ASE2 `F176=F113+F130-L113+F90-L90` (**5 términos**); ASE3 `F316=F238+F254+F217-L217-L238`; ASE4 `F437=F391+F417+F357-L357-L391`; ASE5 `F519=F513+F498+F478-L478-L498`. EXTEMP: ASE1 `F48=F30+F10-L10`, ASE3 `F318=F243+F223-L223`, ASE4 `F439=F401+F369-L369`, pero ASE2 `F178` y ASE5 `F521` son **valor 0 sin fórmula** | Dump `<c>` raw de sheet10.xml | **Mata cualquier estrategia de "offset aritmético"**. R1 exige mapa explícito por ASE + discovery T0 |
| `D104=D9+D28+D47+D66+D85` (shared `D104:F106`, maestro + réplicas por fila), `D105=D10+D29+D48+D67+D86` (verificado), `D109=SUM(D104:D108)` | Dump OuterXml shared formulas | D104:D108 y D109 son fórmulas protegidas; jamás se escriben |
| `D85='AJUSTES - SF-T'!D47` (celda con fórmula, valor 0 en Q1) | Dump sheet18.xml | AJUSTES-SF-T Q1 = 0; no se toca (2.5 es otra HU) |

**NO verificado (y por eso el plan incluye discovery T0 obligatorio, §4 Fase 0):**

1. Para cada bloque ASE de R1/R2/R4: **qué celdas operando son editables (valor) vs derivadas (fórmula)**. Solo se dumpearon las fórmulas visibles, no el carácter de cada operando (ej. ¿F113/F130/L113/F90/L90 son valores editables?).
2. Origen en **fuentes** de cada operando de bloque (¿el leaf reader actual encuentra esos slots por labels? Los labels de bloque Lime/Ciudad/etc. pueden diferir de Promoambiental).
3. Naturaleza de `R1!F178` y `R1!F521` (valor 0 sin fórmula): ¿input manual del proceso manual que debe escribirse como valor, o celda que debe dejarse intacta? El golden Q1 dice D48=0 y D51=0; las fuentes R1 de Lime/Área Limpia deben decir si hay extemporáneo real.
4. ASE4 (Bogotá Limpia) sin columna Especiales en **fuente** R2 — ya manejado (`ServEspK=0`), pero debe re-verificarse que el bloque R4/R2 del template para ASE4 no espera otra cosa.

> **Regla de hierro del plan:** ninguna dirección de celda del mapa multi-ASE entra al código sin pasar por T0. El plan prescribe el **protocolo** de descubrimiento, no direcciones inventadas. Lo ya verificado arriba (visibles + fórmulas CONSOLIDADO) sí es contratable desde el día uno.

### 0.2 Mapeo al Rector (in vs out)

**Entra porque §4 Fase 2 it. 2.1 / §9 / §10 lo piden ahora:**

| Rector | Qué cubre HU-07 |
|---|---|
| §9 it. 2.1 | Iterar las 5 carpetas ASE del período en un solo proceso |
| §4 it. 2.1 + §10 CA-1/CA-2 | Leer R1/R2/R4 de los 5 ASE + cálculo por ASE (GranTotal = suma) |
| §10 CA-3 | CONSOLIDADO D9:D13, D28:D32, D47:D51, D66:D70, D104:D109 correctos para los 5 ASE (vía leaf por bloque + Capa A por ASE) |
| §10 CA-4/CA-5/CA-6/CA-7 | Fórmulas intactas (mapa ampliado), validación multi-ASE, Serilog por ASE, UI modo 5 ASE |

**Sale porque §9 Fase 2 lo asigna a 2.2–2.7 (HUs posteriores):**

- 2.2 conciliación por empresa (hojas EAAB/ENEL/ENERBIT/OCCIDENTE/EAAB-CL, VALIDACION_*).
- 2.3 REPORTE RECAUDO x BANCO. 2.4 BCE SC POR FACT.
- 2.5 AJUSTES-SF-T (SALDOS POR NOTA + RETRIBUCION NEGATIVA; Q1 sigue con AjustesSfT=0; quincena 2 sigue bloqueada en `CalculoRemuneracion`).
- 2.6 DetRetri / DetValiRetri (enteros como objetivo de escritura).
- 2.7 validaciones cruzadas (VALIDACION_TOTAL, VALIDACION_RECIP, etc.).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **Una sola llamada de escritura** por proceso multi-ASE (un workbook de salida, 5 bloques). Ver §2.3. |
| G2 | **Mapa de celdas por ASE explícito** (`WorkbookLeafCellMapPorAse`), no aritmética de offsets. R1 lo exige (fórmulas heterogéneas). Ver §2.1. |
| G3 | **Discovery T0 bloquea el cell-map**: sin dump operando-por-operando no hay código de escritura multi-ASE. Ver §4 Fase 0. |
| G4 | **Orquestador de período en Core** reutilizando el path single-ASE por ASE; fail-fast; sin salida certificada ante cualquier fallo. Ver §2.4. |
| G5 | **Validador endurecido**: N consolidados con matcheo estricto por `Ase.Id` (se elimina el fallback a `FirstOrDefault`), GranTotal = suma, AjustesSfT=0 Q1, gate leaf-vs-consolidado por ASE. Ver §2.5. |
| G6 | **UI delta mínimo**: modo "procesar los 5 ASE" vs ASE seleccionado; sin restyle. Ver §2.6. |
| G7 | **Golden Capa A × 5 ASE** con la misma honestidad HU-06 (leaf salida vs leaf golden; visibles de dominio vs cache golden; nunca cache de salida vs golden). Ver §2.7. |
| G8 | No se reabre semántica HU-04/HU-05/HU-06: `IPlantillaWriter` validation-only intacto; coherencia R1 (`F25` vs Extemp; prohibido TotOpt vs F46) intacta **para ASE1** y extendida por bloque según T0. |

---

## 1. PROPOSE

### 1.1 Intent

Escalar la ruta leaf certificada (leer → calcular → leaf → validar → `GenerarWorkbook`) de 1 ASE a los 5 ASE del período, produciendo **un** workbook de salida con los 5 bloques diligenciados y el CONSOLIDADO completo, con la misma garantía de preservación de fórmulas y la misma honestidad de golden.

### 1.2 In Scope

- Discovery T0 del layout por bloque (protocolo §4 Fase 0) + `WorkbookLeafCellMapPorAse` congelado solo con evidencia.
- Lectura leaf por ASE (`IWorkbookLeafInputReader` ya es por-ASE; verificar labels por bloque, sin cambiar contrato).
- Cálculo multi-ASE (`CalcularConsolidado` ya acepta lista — verificar GranTotal=suma y duplicados; sin cambiar contrato salvo bug).
- Escritura multi-ASE en una sola llamada (overload/nuevo método; un `File.Copy` + una validación pre/post).
- `IValidador` multi-ASE + endurecimiento del matcheo por `Ase.Id` (rompe el fallback `FirstOrDefault` en `ValidadorBasico` y `WorkbookLeafCoherence.ValidarContraResultado`).
- Orquestador de período (`SolicitudProcesoPeriodo` / `ResultadoProcesoPeriodo` / `IProcesadorPeriodo`) + UI modo 5 ASE.
- Golden Capa A extendida a 5 ASE + tests de integración con insumos Q1 reales (5 carpetas).

### 1.3 Out of Scope

Todo §0.2 (2.2–2.7). Además: quincena 2 (sigue lanzando `CalculoInvalidoException`); motor Excel/COM en CI; restyle UI; framework DI; reescritura del path single-ASE (se reutiliza, no se duplica).

### 1.4 Resultado de negocio

El Ingeniero elige período 202607-1, carpeta `REMUNERACION 2026071`, plantilla, carpeta salida y modo "5 ASE"; pulsa Ejecutar; obtiene `Remuneración 202607-1 Total.xlsx` con D9:D13, D28:D32, D47:D51, D66:D70, D104:D109 correctos post-Excel; el log audita por ASE; las pruebas demuestran leaf+visibles por ASE contra el golden Q1.

---

## 2. DESIGN

### 2.1 Mapa 5-ASE verificado (contratable desde el día uno)

Fórmulas CONSOLIDADO por fila (dump sheet18.xml — **evidencia, no invención**):

| Fila CONSOLIDADO | ASE1 | ASE2 | ASE3 | ASE4 | ASE5 |
|---|---|---|---|---|---|
| TOT_OPT D9:D13 | R1!F46 | R1!F176 | R1!F316 | R1!F437 | R1!F519 |
| R2 D28:D32 | R2!E41 | R2!E135 | R2!E247 | R2!E343 | R2!E413 |
| EXTEMP D47:D51 | R1!F48 | R1!F178 (valor, sin fórmula) | R1!F318 | R1!F439 | R1!F521 (valor, sin fórmula) |
| R4 D66:D70 | R4!D67 | R4!D161 | R4!D198 | R4!D312 | R4!D347 |
| Totales D104:D108 | D9+D28+D47+D66+D85 (+ shared por fila) | idem fila 105 | … | … | … |
| GranTotal D109 | SUM(D104:D108) | — | — | — | — |

Fórmulas visibles por bloque (dump raw `<c>` — **evidencia**):

| Bloque | R1 TOT_OPT | R1 EXTEMP | R2 | R4 |
|---|---|---|---|---|
| ASE1 Promoambiental | `F25+F41-L25` (F46) | `F30+F10-L10` (F48) | `E15+E26-K15` (E41) | `D9-P9` (D67) |
| ASE2 Lime | `F113+F130-L113+F90-L90` (F176) | **valor 0, sin fórmula** (F178) | `E85+E103-K85` (E135) | `D98-P98` (D161) |
| ASE3 Ciudad Limpia | `F238+F254+F217-L217-L238` (F316) | `F243+F223-L223` (F318) | `E167+E178-K167` (E247) | `D193-P193` (D198) |
| ASE4 Bogotá Limpia | `F391+F417+F357-L357-L391` (F437) | `F401+F369-L369` (F439) | `E290+E308-K290` (E343) | `D236-P236` (D312) |
| ASE5 Área Limpia | `F513+F498+F478-L478-L498` (F519) | **valor 0, sin fórmula** (F521) | `E385+E374-K374` (E413) | `D344-P344` (D347) |

Valores golden Q1 esperados post-Excel (cache del golden — referencia para A2):

| ASE | D TOT_OPT | D R2 | D EXTEMP | D R4 |
|---|---|---|---|---|
| 1 Promoambiental | 16704332434.57 | 54216385.68 | 11673020 | -12054255.65 |
| 2 Lime | 12157780441.19 | 79400801.26 | 0 | -9889189.72 |
| 3 Ciudad Limpia | 10101988514.82 | 31111803.75 | 10198723.07 | -16103442.89 |
| 4 Bogotá Limpia | 10552409503.83 | 17236000.33 | 5419780 | -21288908.57 |
| 5 Área Limpia | 8519310329.28 | 30774154.23 | 0 | -6421274.68 |

### 2.2 Decisiones de arquitectura

| ID | Opción elegida | Descartada | Por qué |
|---|---|---|---|
| D1 | `WorkbookLeafCellMapPorAse`: mapa explícito por `Ase.Id` (visibles + leaf editables por bloque, congelado tras T0) | Offset aritmético (bloque2 = bloque1 + N filas) | R1 es heterogéneo (3 vs 5 términos; F178/F521 sin fórmula). La aritmética falla en ASE2/ASE3. Riesgo §11.3 del Rector (no hardcodear posiciones) |
| D2 | Overload `GenerarWorkbook(origen, salida, resultado, IReadOnlyList<WorkbookLeafInputs>)` (o método `GenerarWorkbookPeriodo`) en `IWorkbookLeafWriter` | 5 llamadas single-ASE (5 copias) o 5 writers | Un solo `File.Copy` + una validación pre/post = una salida atómica; 5 copias romperían atomicidad y el hash A4 |
| D3 | Orquestador `IProcesadorPeriodo` en Core que **reutiliza** `IProcesadorRemuneracion` por ASE hasta validación, y escribe una sola vez | Duplicar el pipeline en el procesador de período | El path single-ASE está certificado; se compone, no se copia. La escritura se difiere al final para atomicidad |
| D4 | Matcheo estricto `Consolidados.Single(c => c.Ase.Id == leaf.Ase.Id)` en validador y coherencia | Mantener fallback `FirstOrDefault` | Con 5 consolidados el fallback es bug silencioso (valida ASE1 contra leaf de ASE3). Endurecer es obligatorio |
| D5 | `Validador`: `Consolidados.Count == 5` en modo período (rango 1..5 con mensaje distinto para 1 vs N, para no romper tests single-ASE existentes salvo actualización explícita) | Exigir exactamente 5 siempre | Los 8 tests actuales single-ASE deben seguir verdes; el modo período exige 5 |
| D6 | UI: CheckBox "Procesar los 5 ASE" (default según lo que decida el Ingeniero en aprobación) + combo ASE se deshabilita en modo 5 | Grid de resultados / selector multi-ASE / restyle | Delta mínimo HU-06 G5; OPA sigue siendo Ejecutar |
| D7 | Golden Capa A parametrizada por ASE (misma matriz HU-06 × 5) | Un golden "promedio" o comparar solo CONSOLIDADO | Cada bloque tiene su verdad; promediar escondería un ASE roto |

### 2.3 Escritura multi-ASE (una sola llamada)

```text
Plantilla origen
  └─► File.Copy → ruta salida (una vez)
        └─► ValidarFormulasProtegidas (mapa ampliado: 5 visibles R1 + 5 R2 + 5 R4 + D9:D13/D28:D32/D47:D51/D66:D70/D104:D109)
              └─► por cada leaf (5): EscribirCeldasLeaf del bloque (solo celdas del mapa por ASE)
                    └─► revalidar fórmulas protegidas → guardar
```

Protecciones: ninguna celda con `<f>` puede ser objetivo de escritura (guard existente `cell.CellFormula is not null` se mantiene — es lo que protege F178/F521 si resultan ser fórmula en otra plantilla; si son valor, T0 decide si se escriben o se dejan). Ante cualquier fallo: borrar salida parcial (patrón existente try/catch+delete). La plantilla origen jamás se muta.

### 2.4 Orquestación de período (Form1 flaco, Core orquesta)

```csharp
public sealed class SolicitudProcesoPeriodo
{
    public Periodo Periodo { get; set; } = new();
    public string CarpetaPeriodo { get; set; } = "";   // REMUNERACION 2026071 (contiene las 5 carpetas)
    public string RutaPlantilla { get; set; } = "";
    public string RutaSalida { get; set; } = "";        // carpeta + Periodo.NombreArchivo (igual que HU-06)
}

public sealed class ResultadoProcesoPeriodo
{
    public ResultadoRemuneracion Resultado { get; set; } = new();          // 5 consolidados
    public IReadOnlyList<WorkbookLeafInputs> Leafs { get; set; } = [];     // 5, uno por ASE
    public string RutaSalida { get; set; } = "";
}

public interface IProcesadorPeriodo
{
    ResultadoProcesoPeriodo Ejecutar(SolicitudProcesoPeriodo solicitud, IProgress<string>? progreso = null);
}
```

Flujo `ProcesadorPeriodo.Ejecutar`: resolver 5 carpetas vía `ArchivoFuenteLocator.ObtenerCarpetasAse` + prefijos `CarpetasAse.Prefijos` (falta alguna → `ArchivoFuenteNoEncontradoException`, sin salida); por cada ASE (orden 1..5): LeerR1/R2/R4 → acumular tupla; `CalcularConsolidado(periodo, 5 tuplas)` (verifica §3 Req 1); `LeerLeafInputs` por ASE; `Validar(resultado, leafs)` multi-ASE (lista no vacía → `CalculoInvalidoException`, sin salida); **una** llamada `GenerarWorkbook(origen, salida, resultado, leafs)`. `IProgress<string>` reporta por ASE para la barra/log. Reutilización máxima del path single-ASE: el procesador de período delega los pasos por ASE en helpers compartidos con `ProcesadorRemuneracion` (extraer métodos internos o composición), no copia el código.

### 2.5 `IValidador` multi-ASE (endurecido, sin reabrir HU-06)

Nuevo overload `Validar(ResultadoRemuneracion, IReadOnlyList<WorkbookLeafInputs>)`. Reglas:

1. `Consolidados.Count` == `leafs.Count`; en modo período se exige 5 (mensaje explícito si falta un ASE).
2. Sin duplicados de `Ase.Id` (defensa en profundidad; el cálculo ya los rechaza).
3. `AjustesSfT == 0` en los 5 (Q1; §4.2 del Rector).
4. Gate leaf-vs-consolidado **por ASE con matcheo estricto** (`Single` por Id; si falta → error, nunca fallback): `leaf.R1.F25`-equivalente vs `Extemp`, R2 vs `R2TotalOportuno`, R4 vs `ReversionR4` — con las equivalencias por bloque que congele T0 (para ASE1 las actuales; prohibido TotOpt vs visible).
5. `GranTotal == Σ TotalAse` (±0.5) + `TotalAse` aritmético por ASE.
6. El validador NO abre `.xlsx` (igual que HU-06).

### 2.6 UI — Visual Design Intent (delta mínimo)

- Densidad Balanced, mismos GroupBoxes. Un `chkCincoAse` ("Procesar los 5 ASE") junto a `cmbAse`; en modo 5 el combo se deshabilita (pero conserva selección para modo single).
- Progreso: máximo = hitos × 5 + escritura (ej. 5×8+2); `ProgressBarStyle.Continuous` existente.
- Resumen en `txtLog`: una línea por ASE con visibles esperados post-Excel (F176/D10, F178/D48, …) + GranTotal; etiquetar siempre como "esperado post-Excel", nunca como cache de salida. Prohibido restyle/colores/iconos.

### 2.7 Golden Capa A × 5 (honestidad HU-06 preservada)

| # | Qué | Contra qué | Tolerancia |
|---|---|---|---|
| A1 | Leaf escritos en la **salida** (celdas del mapa por ASE) | Mismas celdas **leaf** del golden | ±0.5 |
| A2 | Visibles de **dominio** por ASE (aritmética con mapa por bloque) | Cache golden de visibles del bloque + CONSOLIDADO D_fila(ASE) | ±0.5 |
| A3 | Fórmulas protegidas del mapa ampliado | Siguen siendo fórmula en la salida | n/a |
| A4 | SHA256 plantilla origen | Igual antes/después | n/a |
| A5 | Prohibido comparar cache de fórmula **de la salida** vs golden; prohibido TotOpt HU-02 vs visibles R1 | — | prohibido |
| A6 (nuevo) | D104:D108 = suma de su fila y D109 = SUM(D104:D108) como **fórmulas** + aritmética de dominio Σ por ASE | Estructura + valores de dominio | ±0.5 valores |

Capa B (manual Excel: abrir, recalcular, comparar D9:D13/D28:D32/D47:D51/D66:D70/D104:D109 vs golden) sigue fuera de CI, protocolo §5.3.

### 2.8 File changes previstos

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapPorAse.cs` | Crear | Mapa explícito por `Ase.Id` (congelado tras T0) |
| `Remuneracion.Core/Interfaces/IProcesadorPeriodo.cs` | Crear | Caso de uso de período |
| `Remuneracion.Core/Models/SolicitudProcesoPeriodo.cs` | Crear | Input período |
| `Remuneracion.Core/Models/ResultadoProcesoPeriodo.cs` | Crear | Output período (5 leafs) |
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` | Crear | Orquestación 5 ASE, una escritura |
| `Remuneracion.Core/Interfaces/IWorkbookLeafWriter.cs` | Modificar | Overload multi-ASE (lista de leafs) |
| `Remuneracion.Core/Interfaces/IValidador.cs` | Modificar | Overload multi-ASE |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Modificar | Reglas §2.5 + matcheo estricto (rompe fallback) |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs` | Modificar | `ValidarContraResultado` multi-ASE, matcheo estricto |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Implementar overload multi-ASE + validación mapa ampliado |
| `Remuneracion.WinForms/Form1.cs` + `Form1.Designer.cs` | Modificar | Modo 5 ASE, progreso ×5, resumen por ASE |
| `Remuneracion.WinForms/Program.cs` | Modificar | Registrar `ProcesadorPeriodo` en composition root |
| `Remuneracion.IntegrationTests/GoldenMultiAseTests.cs` | Crear | Capa A × 5 (§2.7) |
| `Remuneracion.IntegrationTests/ValidadorMultiAseTests.cs` | Crear | Dominio multi-ASE sin Excel |
| `Remuneracion.IntegrationTests/ProcesadorPeriodoTests.cs` | Crear | Integración 5 carpetas reales, salida temp |

**No tocar (salvo bug blocker):** `IPlantillaWriter` y modos validation-only (HU-04); contratos `IRecaudoReader`/`IWorkbookLeafInputReader`/`ICalculoRemuneracion`; semántica single-ASE de `ProcesadorRemuneracion` (se reutiliza); `requirements/` legado.

---

## 3. SPEC / Acceptance Criteria

### Requirement 1 — Cálculo multi-ASE

El sistema **MUST** calcular 5 consolidados (uno por ASE, sin duplicados) con `GranTotal = Σ TotalAse`.

- GIVEN las 5 carpetas Q1 con R1/R2/R4 presentes → WHEN `CalcularConsolidado(periodo, 5 tuplas)` → THEN 5 consolidados, `GranTotal` = suma ±0.5.
- GIVEN tuplas con `Ase.Id` duplicado → THEN `CalculoInvalidoException` (ya existe; re-verificar con test).
- GIVEN quincena 2 → THEN sigue bloqueado (sin cambio).

### Requirement 2 — Lectura leaf por ASE

El sistema **MUST** producir un `WorkbookLeafInputs` por ASE desde sus 3 fuentes, con visibles esperados que cuadren con el bloque golden correspondiente (±0.5). **MUST NOT** asumir que los labels/slots de ASE1 sirven para ASE2..5: T0 define el mapeo por bloque; si un bloque no se puede mapear por labels, el proceso falla con mensaje que identifica el ASE.

- GIVEN carpeta `2-Lime` → WHEN `LeerLeafInputs` → THEN aritmética del bloque (F113+F130-L113+F90-L90, E85+E103-K85, D98-P98) = golden Lime ±0.5. Idem ASE3/4/5 con sus fórmulas §2.1.

### Requirement 3 — Escritura única multi-ASE preservando fórmulas

El sistema **MUST** generar el workbook con una sola copia y escribir solo celdas del mapa por ASE. **MUST NOT** escribir ninguna celda con fórmula; **MUST NOT** mutar la plantilla; ante fallo **MUST** borrar la salida parcial.

- GIVEN 5 leafs válidos → WHEN overload multi-ASE → THEN cambian solo celdas del mapa; visibles R1/R2/R4 por bloque + D9:D13/D28:D32/D47:D51/D66:D70/D104:D109 siguen siendo fórmula.
- GIVEN gate de coherencia roto en cualquier ASE → THEN no hay archivo certificado.

### Requirement 4 — Validación multi-ASE endurecida

`IValidador` **MUST** exigir 5 consolidados + 5 leafs matcheados por `Ase.Id` estricto, AjustesSfT=0, gates por ASE, GranTotal=Σ. **MUST NOT** usar fallback a primer consolidado; **MUST NOT** comparar TotOpt HU-02 vs visibles R1; **MUST NOT** abrir Excel.

### Requirement 5 — Orquestación de período fail-fast

El sistema **MUST** procesar los 5 ASE en orden 1..5 con log por ASE y **MUST NOT** certificar salida si falta una carpeta/fuente, falla validación o falla escritura.

- GIVEN falta R4 en ASE4 → WHEN Ejecutar período → THEN `ArchivoFuenteNoEncontradoException` que nombra el ASE, sin archivo de salida.
- GIVEN validación rota en ASE3 → THEN `CalculoInvalidoException`, sin archivo de salida.

### Requirement 6 — UI modo 5 ASE honesta

La UI **MUST** ofrecer modo 5 ASE vs ASE seleccionado con el mismo patrón de validación de rutas (fuentes+plantilla+salida distinta, overwrite con MessageBox). El resumen **MUST** listar visibles esperados post-Excel por ASE + GranTotal, sin presentar agregados HU-02 como valores CONSOLIDADO.

### Requirement 7 — Golden Capa A × 5

Las pruebas **MUST** ejecutar la matriz §2.7 para los 5 ASE con insumos Q1 reales. **MUST NOT** comparar cache de fórmula de la salida vs golden.

### AC resumidos vs Propuesta §10

| CA | HU-07 |
|---|---|
| CA-1 | Lee R1/R2/R4 de los 5 ASE (fail-fast nombra el ASE) |
| CA-2 | Cálculo + visibles de dominio por ASE (TotOpt fuente ≠ visible R1, como HU-05/06) |
| CA-3 | Capa A × 5; cache CONSOLIDADO post-Excel = Capa B manual residual |
| CA-4 | Reassert mapa ampliado + plantilla no mutada (hash) |
| CA-5 | `IValidador` multi-ASE + coherencia por bloque |
| CA-6 | Serilog + `txtLog` por ASE y por paso |
| CA-7 | Período, carpeta, plantilla, salida + modo 5 ASE |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 700–1100 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 discovery T0 + mapa → PR2 validador+coherencia endurecidos → PR3 writer overload → PR4 procesador período + UI → PR5 golden × 5 |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes. Chained PRs recommended: Yes. Chain strategy: pending.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 0 | Discovery T0 + mapa por ASE congelado | PR 1 | Bloquea todo lo demás; solo lectura + datos, sin cambiar escritura |
| 1 | Validador + coherencia multi-ASE (matcheo estricto) | PR 2 | Depende de PR 1 (equivalencias por bloque) |
| 2 | Writer overload multi-ASE + mapa ampliado | PR 3 | Depende de PR 1 |
| 3 | `ProcesadorPeriodo` + UI modo 5 ASE | PR 4 | Depende de PR 2 y PR 3 |
| 4 | Golden × 5 + integración período | PR 5 | Depende de PR 3 y PR 4; 8 tests previos siempre verdes |

### Phase 0 — Discovery T0 (BLOQUEANTE, solo lectura, sin cambiar escritura)

- [ ] 0.1 Dumpear por OpenXML (mismo método de esta planificación: unzip + `<c>` raw) **todas** las celdas operando de cada fórmula visible §2.1 y clasificarlas: valor editable vs fórmula. Ej. R1 ASE2: F113, F130, L113, F90, L90; ASE3: F238, F254, F217, L217, L238, F243, F223, L223; ASE4: F391, F417, F357, L357, L391, F401, F369, L369; ASE5: F513, F498, F478, L478, L498; R2: E85/E103/K85, E167/E178/K167, E290/E308/K290, E385/E374/K374; R4: D98/P98, D193/P193, D236/P236, D344/P344. Registrar `<f>` (si existe) y `<v>` de cada una.
- [ ] 0.2 Resolver naturaleza de `R1!F178` y `R1!F521` (valor 0 sin fórmula): cruzar con fuentes R1 de Lime y Área Limpia (¿hay extemporáneo real que deba escribirse como valor, o el 0 manual se deja intacto?). Decisión escrita en el plan antes de codificar.
- [ ] 0.3 Verificar que `IWorkbookLeafInputReader` localiza por labels los slots de cada bloque en las **fuentes** de ASE2..5 (las fuentes son por ASE; los labels de bloque del template no tienen por qué coincidir). Incluye re-verificación ASE4-sin-Especiales en fuente y bloque.
- [ ] 0.4 Congelar `WorkbookLeafCellMapPorAse` (visibles + editables por `Ase.Id` 1..5) con tabla de evidencia celda-por-celda. **Nada del mapa entra al código sin esta tabla.**
- [ ] 0.5 Confirmar filas de totales por ASE en el golden (D104:D108 aritmética de dominio vs cache) y que `D85:D89`/`'AJUSTES - SF-T'!D47` = 0 en Q1.

### Phase 1 — Dominio (validador + coherencia)

- [ ] 1.1 Overload `IValidador.Validar(ResultadoRemuneracion, IReadOnlyList<WorkbookLeafInputs>)` + reglas §2.5 (5 en modo período; matcheo `Single` por Id; AjustesSfT=0; gates por bloque T0; GranTotal=Σ).
- [ ] 1.2 Endurecer `Validar(resultado, leaf)` single-ASE: `Single` por Id (rompe fallback `FirstOrDefault`; actualizar tests que dependan del fallback).
- [ ] 1.3 `WorkbookLeafCoherence.ValidarContraResultado` multi-ASE (itera leafs, matcheo estricto) + mantener validación single-ASE existente.
- [ ] 1.4 Re-verificar `CalcularConsolidado` con 5 tuplas: GranTotal=Σ, duplicados fallan (tests, sin cambio de contrato salvo bug).

### Phase 2 — Escritura multi-ASE

- [ ] 2.1 Overload `IWorkbookLeafWriter.GenerarWorkbook(origen, salida, resultado, IReadOnlyList<WorkbookLeafInputs>)`: una copia, validación pre/post con mapa ampliado, escritura por bloque, borrado de parcial ante fallo.
- [ ] 2.2 `WorkbookLeafCellMapPorAse.cs` (datos de T0-0.4) + validación de que ningún objetivo tiene `<f>` al momento de escribir.
- [ ] 2.3 Proteger D104:D108/D109 y todas las filas CONSOLIDADO por ASE en `ProtectedFormulas` ampliado (incluye shared-formula awareness: validar maestro + réplicas).

### Phase 3 — Orquestación + UI

- [ ] 3.1 `SolicitudProcesoPeriodo`, `ResultadoProcesoPeriodo`, `IProcesadorPeriodo`, `ProcesadorPeriodo` (§2.4): 5 carpetas vía `ArchivoFuenteLocator` + `CarpetasAse.Prefijos`, reutilización del path por ASE, fail-fast, una escritura, `IProgress<string>` por ASE.
- [ ] 3.2 `Program.cs`: composition root suma `ProcesadorPeriodo`.
- [ ] 3.3 `Form1`: `chkCincoAse`, progreso ×5, resumen por ASE honesto post-Excel + GranTotal; mismos patrones de validación/overwrite/log dual Serilog.

### Phase 4 — Pruebas y evidencia

- [ ] 4.1 `ValidadorMultiAseTests` (in-memory): happy path 5, mismatch por ASE (nombra el ASE), Ajustes≠0, GranTotal roto, duplicados, fallback eliminado.
- [ ] 4.2 `ProcesadorPeriodoTests`: 5 carpetas Q1 reales, salida temp; caso falta-fuente en un ASE (sin salida).
- [ ] 4.3 `GoldenMultiAseTests`: matriz §2.7 × 5 ASE; fixtures golden `Remuneracion 202607-1 Total.xlsx` + `REMUNERACION 2026071/{1..5}-*/` (carpeta `5-Área Limpia` con tilde — cuidado encoding en tests/locator).
- [ ] 4.4 Los 8 tests existentes siguen verdes; build 0 warnings; CRLF; sin commit (regla del proyecto).

### Phase 5 — Documental

- [ ] 5.1 Capa B manual §5.3 ejecutada al menos una vez y evidenciada (captura/nota), sin fingirla como gate de merge.
- [ ] 5.2 Cierre deja explícito qué queda para 2.2–2.7 (siguiente HU propuesta: 2.2 conciliación por empresa).

---

## 5. Test Strategy

| Capa | Qué | Cómo |
|---|---|---|
| Unidad | `ValidadorBasico` multi-ASE + matcheo estricto | Modelos in-memory, sin Excel |
| Unidad | `CalcularConsolidado` 5 tuplas (suma, duplicados) | In-memory |
| Integración | `ProcesadorPeriodo` 5 carpetas reales | Insumos Q1, salida temp, fail-fast verificado |
| Golden Capa A × 5 | Matriz §2.7 por ASE | OpenXML read-only + aritmética de dominio; plantilla=golden como en HU-06 (A5 aplica) |
| UI | Modo 5 ASE | Funcional manual (sin harness WinForms) |
| Capa B | CA-3 post-Excel | Manual — §5.3 |

### 5.1 Fixtures

- Golden/plantilla estructura: `Docs/Insumos/Remuneracion 202607-1 Total.xlsx`.
- Fuentes: `Docs/Insumos/REMUNERACION 2026071/{1-Promoambiental,2-Lime,3-Ciudad Limpia,4-Bogota Limpia,5-Área Limpia}/` (prefijos R1/R2/R4 §2.2 HU-06; ASE4 sin Especiales).
- Valores esperados por ASE: tabla §2.1 (cache golden, tolerancia ±0.5).

### 5.2 Casos negativos obligatorios

Falta carpeta de un ASE; falta R4 en ASE4; gate leaf roto solo en ASE3 (el error nombra ASE3); `F178`/`F521` según decida T0-0.2; salida == plantilla (no in-place); overwrite cancelado.

### 5.3 Protocolo manual Capa B (no CI)

1. Generar salida a ruta distinta del golden. 2. Abrir en Excel, recalcular. 3. Comparar D9:D13, D28:D32, D47:D51, D66:D70, D104:D108, D109 vs golden ±0.5 (esperados §2.1). 4. Evidencia adjunta. No es gate de merge.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | Período = orquestación; procesador single-ASE = pasos por ASE; validador = reglas; writer = I/O; mapa por ASE = datos. |
| **O** | Se agregan overloads/contratos nuevos; el path single-ASE certificado no se reescribe. |
| **L** | `OpenXmlPlantillaWriter` suma overload multi-ASE sin cambiar el comportamiento single-ASE ni validation-only. |
| **I** | `IProcesadorPeriodo` separado de `IProcesadorRemuneracion`; validador/writer crecen por overload, no por ruptura. |
| **D** | Core define `IProcesadorPeriodo`/overloads; Infrastructure/WinForms componen. Locator sigue en Infrastructure. |

### 6.2 Best Practices

- La verdad del workbook manda: mapa explícito por ASE verificado, no offsets asumidos (Rector §11.3–11.4).
- Una escritura atómica; fórmulas intactas; plantilla nunca mutada; hash A4.
- Golden honesto por ASE (OpenXML no recalcula; §2.7 A5).
- Fail-fast nombra el ASE; sin salida certificada ante fallo.
- Serilog + `txtLog` por ASE (Rector §6).

### 6.3 Performance

- 5 ASE × (3 lecturas + 1 leaf) + 1 copia + ~55 escrituras leaf + revalidación acotada. Irrelevante a esta escala; `Task.Run` existente para no congelar el form.

**Veredicto:** APROBADO como it. 2.1 de Fase 2 **si** T0 congela el mapa con evidencia y se acepta CA-3 parcial (Capa A en CI, Capa B manual).

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| Algún operando de bloque R1 es fórmula (no editable) y rompe el modelo leaf | Alta | T0-0.1 lo detecta antes de codificar; el plan no prescribe direcciones sin evidencia |
| `F178`/`F521` requieren semántica de escritura no prevista | Media | T0-0.2 decide por escrito; si son inputs manuales, el reader los deriva de fuente o se dejan intactos con test que lo fija |
| Labels de fuentes ASE2..5 difieren y el leaf reader no los encuentra | Media | T0-0.3; el fallo nombra el ASE (fail-fast), no se finge mapeo |
| Alguien "cierra" CA-3 comparando cache de salida vs golden (plantilla=golden) | Alta | A5/A6 en tests; este plan lo prohíbe |
| Fallback `FirstOrDefault` enmascara un cruce de ASE | Alta | Matcheo `Single` + tests negativos por ASE |
| Inflar a 2.2–2.7 dentro de esta HU | Media | §0.2 out-of-scope explícito; rechazar PRs que lo metan |
| Carpeta `5-Área Limpia` (tilde/encoding) rompe locator/tests en Windows | Media | Tests usan nombres reales; `StartsWith(id)` existente + caso de prueba |
| Recálculo Excel nunca se hace | Media | Capa B residual explícita §5.3 |

---

## 8. Rollback

- Revertir PRs en orden inverso (golden → UI/período → writer → validador → mapa/T0).
- HU-04/HU-05/HU-06 intactas sin esta HU: el modo single-ASE sigue operativo.
- `Docs/Insumos/` untracked: jamás usar esos xlsx como destino; cualquier pisada se restaura desde backup.
- Si T0 demuestra que un bloque no es mapeable por labels, la HU se recorta a los ASE mapeables con plan rebaseado (no se inventa mapeo).

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §4 it. 2.1 / §6 / §9 Fase 2 / §10 (±0.5). 2.2–2.7 quedan fuera, citados a §9.
2. Layout 5-ASE §2.1 es evidencia verificada (dump OpenXML); el mapa editable por bloque queda pendiente de T0.
3. Una sola escritura atómica por período; R1 heterogéneo prohíbe offsets aritméticos.
4. Matcheo estricto por `Ase.Id` (se elimina el fallback silencioso).
5. Golden Capa A × 5 con honestidad HU-06; Capa B manual residual.
6. UI delta mínimo; OPA = Ejecutar; sin restyle.
7. Apply espera aprobación + estrategia de PRs encadenados (Unidad 0–4, §4).

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
