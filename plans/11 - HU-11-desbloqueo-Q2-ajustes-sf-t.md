# Plan 11 — HU-11: Desbloqueo Q2 + AJUSTES-SF-T (Fase 2, it. 2.5)

> **Historia:** abrir el modelo de quincena 2 (reemplazar el fail-fast por `AjustesSfT` real desde SALDOS POR NOTA + RETRIBUCION NEGATIVA) con regresion Q1 intacta, dentro de la misma escritura atomica multi-ASE certificada por HU-07/HU-08/HU-09/HU-10.
> **Rector (IRRENUNCIABLE):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 2 it. **2.5 AJUSTES-SF-T**, §6 arquitectura (capas, preservacion de formulas, Serilog, lectura directa de fuente), §10 CA + tolerancia ±0.5. **Entra SOLO 2.5.** Salen 2.6 DetRetri, 2.7 validaciones cruzadas — son HUs posteriores.
> **Origen funcional:** `Detalle de plantilla.docx` (Inst. hojas `AJUSTES - SF-T`, `SALDOS POR NOTA`, `RETRIBUCION NEGATIVA` → R5 CONSOLIDADO) y `Prompt Maestro Vo.docx` (paso de pegado de saldos por nota y retribucion negativa). `Proceso de Recaudo.docx` **NO aplica** a esta HU. Citados por seccion en §1.5/§3.
> **Continuidad:** HU-01..HU-10 cerradas (ruta leaf multi-ASE atomica single-write, mapas explicitos por `Ase.Id`, validador estricto, UI modo 5-ASE, Golden Capa A, build 0/0, tests 78/78, harness 24/24). Este plan NO reabre su semantica.
> **Estado actual:** Q2 hoy hace fail-fast con `CalculoInvalidoException` (`CalculoRemuneracion.CalcularConsolidado`, Q2 sin modelar); en Q1 `AjustesSfT = 0`.
> **Estado:** PENDIENTE DE APROBACION — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-08

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. La unica incognita real —la composicion exacta de D85:D89— **no se asume**: la dicta el T0 contra el golden (§4 Fase 0, tarea 0.3), con doble desenlace soportado sin rebase (ver D2). El apply espera aprobacion explicita.

### 0.1 Que esta verificado y que NO (honestidad tecnica)

**Verificado en esta planificacion** (evidencia del orquestador + lectura de codigo y disco propias):

| # | Hecho | Metodo | Consecuencia |
|---|---|---|---|
| V1 | Q2 fail-fast vive en `CalculoRemuneracion.CalcularConsolidado` (lineas 83-87); `Calcular` fija `AjustesSfT = 0m` (linea 41); `ConsolidadoAse.TotalAse` ya suma `AjustesSfT` (getter, sin cambio necesario) | Lectura de codigo | El desbloqueo es **agregar overloads con ajustes**, no reescribir el motor; el path Q1 queda intacto por construccion |
| V2 | `ArchivoFuenteLocator.BuscarArchivo` es prefix-based case-insensitive (sin rango de fechas en el match) | Lectura de codigo | El locator por prefijo **ya es agnostico al rango**; los finders nuevos solo suman prefijos, nunca fechas |
| V3 | Insumos Q2 5/5: cada carpeta `1-Promoambiental … 5-Area Limpia` trae `SaldosaFavorAplicadosPorNotas_to_date01072026…_to_date31072026…` y `RetribucionNegativa_to_date…`; mayoria rango 01072026–31072026, pero **ASE4-RetribucionNegativa trae 16072026–31072026** | `Get-ChildItem` Insumos | Confirma V2: el rango NO es homogeneo; prohibido codificar fechas en el locator (test negativo dedicado) |
| V4 | ASE5-Q2 **solo trae Balance Optimizado** (`R4-BalanceSubsidioyContribuciones-Optimizado_…`, sin variante plana) | `Get-ChildItem` Insumos | El fallback `BuscarBalance` (base → Optimizado) pasa de opcional a **ejercitado obligatoriamente** en Q2; test dedicado OBLIGATORIO |
| V5 | Hay **DOS plantillas 202607-2**: `Plantilla  _ Remuneracion 202607-2 Total.xlsx` y `Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx` | `Get-ChildItem` Insumos | T0-0.1 fija cual es el **golden canonico** (la otra queda como estructura/control, nunca oraculo de merge) |
| V6 | Docs base: D85:D89 = AJUSTES-SF-T (formulada de SALDOS POR NOTA + RETRIBUCION NEGATIVA → R5 CONSOLIDADO); INTERVENTORIA L25:N31 por ASE; columna I "Especiales" siempre en plantilla, a veces ausente en fuente (mapeo por titulo) | Documentos base (orquestador) | Base funcional del mapa T0; INTERVENTORIA entra al mapa **protegido**, no al escribible |
| V7 | `ValidadorBasico` exige `AjustesSfT == 0` en Q1 en las tres sobrecargas (single, leaf, multi-ASE) + `GranTotal == Σ TotalAse` ya incluye `AjustesSfT` via el getter | Lectura de codigo | Regresion Q1 = no tocar esas reglas; Q2 suma reglas nuevas por quincena, no las reemplaza |
| V8 | `BuscarBalance` ya implementa base-primero + fallback Optimizado (HU-10 V8); `ILocalizadorArchivosAse` lo documenta | Lectura de codigo | 2.5 = **extension del patron HU-07..HU-10**, no arquitectura nueva |
| V9 | Nombres en disco con diacriticos (`RetribucionNegativa`, `ReversionPorComponente`) se listan con encoding degradado (`Retribuci�nNegativa`) segun herramienta | `Get-ChildItem` | El matcher del locator nuevo debe ser **insensible a diacriticos** (normalizacion) o usar stem sin acento; T0-0.4 lo verifica a nivel bytes |

**NO verificado (y por eso T0 es bloqueante, §4 Fase 0):**

1. Composicion exacta de D85:D89 por fila (¿`'AJUSTES - SF-T'!D47..D51` directo como en Q1, u otra formula con R5?).
2. Caracter valor-vs-formula celda por celda de la cadena AJUSTES-SF-T: hoja `AJUSTES - SF-T` (visibles D47..D51 por ASE), bloques SALDOS POR NOTA y RETRIBUCION NEGATIVA (¿hojas propias o secciones? ¿operandos editables?).
3. Layout de headers de las dos fuentes nuevas por ASE (¿columna "Especiales" presente/ausente por ASE?; ¿valores negativos por componente en Retribucion en que columna?).
4. Ubicacion/etiqueta exacta de totales en ambas fuentes (¿"Total General"? ¿"Grand Total"? ¿otra?).
5. Cual de las dos plantillas 202607-2 es el golden canonico.
6. Paridad de layouts Q2-vs-Q1 en los demas reportes (R1/R2/R4/Banco/Balance) — si alguno diverge, recorte honesto, no invencion.
7. Si alguna celda objetivo de 2.5 depende de un externalLink roto.

> **Regla de hierro del plan:** ninguna direccion de celda de 2.5 entra al codigo sin pasar por T0, y la composicion D85:D89 no se codifica por hipotesis sino por prueba contra el golden. Lo ya verificado arriba (V1–V9) si es contratable desde el dia uno.

### 0.2 Mapeo al Rector (in vs out)

**Entra porque §9 it. 2.5 / §6 / §10 lo piden ahora:**

| Rector | Que cubre HU-11 |
|---|---|
| §9 it. 2.5 | AJUSTES-SF-T: SALDOS POR NOTA + RETRIBUCION NEGATIVA → `AjustesSfT` por ASE (solo Q2) |
| §5 / §7.1 | Lectura directa de ambas fuentes por ASE + pegado en valores segun mapa T0 |
| §10 CA-1/CA-2 | Leer saldos-nota y retribucion de los 5 ASE + coherencia `AjustesSfT` = fuente ±0.5 |
| §10 CA-3/CA-4 | D85:D89 correctos post-Excel (Capa A + Capa B residual); formulas intactas (mapa ampliado 2.5) |
| §10 CA-5/CA-6/CA-7 | Validacion Q2 por ASE, Serilog por ASE, UI sin cambios salvo resumen |

**Sale porque §9 lo asigna a 2.6–2.7 / Fase 3 (HUs posteriores) — EXPLICITO:**

- 2.6 `DetRetri2026072` / `DetValiRetri2026072` como objetivo de escritura (sus formulas se protegen; su logica no se implementa).
- 2.7 validaciones cruzadas (`VALIDACION_TOTAL`, `VALIDACION_RECIP/ENEL/…`, bloque AJUSTES 16+ mas alla de protegerlo). Se protegen sus formulas; su logica no se implementa.
- `INTERVENTORIA` (L25:N31 por ASE — entra al mapa protegido, no se toca), `ANT EXT-REV` (protegida, no se toca).
- **Insert/delete de filas:** prohibidos. Si la plantilla no alcanza para el detalle Q2, **fail-fast honesto como en HU-07** (error que nombra ASE + reporte, sin salida certificada).
- Reescritura del path HU-07..HU-10 (se extiende, no se duplica); restyle UI; DI framework.

### 0.3 Decisiones ejecutivas (aprobacion = aceptar estas)

| ID | Decision |
|---|---|
| G1 | **T0 bloqueante con protocolo de composicion D85:D89** (§2.2). Sin dump de la cadena AJUSTES-SF-T no se escribe ni una celda 2.5. |
| G2 | **AJUSTES-SF-T como hoja FORMULADA validada, no escrita** (doctrina BCE-H de HU-10, ver D2): los visibles D47..D51 entran al mapa protegido; el writer solo toca operandos editables que T0 demuestre — si no hay ninguno, HU-11 = calculo + validacion + golden sin escritura nueva, sin rebase. |
| G3 | **Q1 intacta por construccion:** overloads nuevos para ajustes; el path Q1 (fail-fast Q2 en overload viejo, `AjustesSfT == 0`, tests Q1) no se toca. |
| G4 | **Lectura header-driven con "Especiales" opcional** (mismo patron R2: `TieneColumnaEspeciales`; ASE4-Bogota Limpia sin Especiales → 0). Mapeo por titulo, nunca por columna fija. |
| G5 | **Locator agnostico al rango y a diacriticos** (V2/V3/V9): prefijos sin fechas; match normalizado sin acentos. Test negativo con el rango ASE4 (1607–3107) dedicado. |
| G6 | **Golden Q2 unico y canonico** fijado por T0-0.1 entre las dos plantillas 202607-2 (la otra = control, nunca oraculo). Validacion D85:D89 ±0.5 + `TotalAse`/`GranTotal` actualizados (ya incluyen `AjustesSfT` via getter). |
| G7 | **Test variante Balance-Optimizado OBLIGATORIO** (V4): ASE5-Q2 solo trae Optimizado; el fallback deja de ser cobertura opcional. |
| G8 | **Una sola llamada de escritura** por proceso (mismo overload lista HU-07..HU-10). Sin segundo pase. |
| G9 | **Golden Capa A extendida a 2.5** con honestidad HU-06..HU-10 (leaf salida vs leaf golden; visibles de **dominio** vs cache golden; nunca cache de salida vs golden). Capa B manual residual. |

---

## 1. PROPOSE

### 1.1 Intent

Abrir el modelo Q2: calcular `AjustesSfT` real por ASE desde SALDOS POR NOTA + RETRIBUCION NEGATIVA, validar la cadena formulada AJUSTES-SF-T sin escribirla, y certificar D85:D89 + totales contra el golden Q2 canonico — sin mover una coma del comportamiento Q1.

### 1.2 In Scope

- Discovery T0 (§4 Fase 0) + `WorkbookLeafCellMapAjustesSfT` congelado solo con evidencia (incluida la composicion D85:D89 y el veredicto editable-vs-formula).
- Modelos `SaldosNotasAseInputs` + `RetribucionNegativaAseInputs` + `AjustesSfTInputs` (patron `TieneColumnaEspeciales` R2) + extension nullable de `WorkbookLeafInputs`.
- Overloads `ICalculoRemuneracion` con ajustes (Q2); path Q1 intacto.
- Readers header-driven de ambas fuentes + finders de locator agnosticos a rango/diacriticos.
- Validacion pre/post con mapa ampliado (AJUSTES-SF-T visibles, D85:D89, INTERVENTORIA L25:N31, ANT EXT-REV, DetRetri Q2 como protegidas).
- Gates Q2 en `IValidador` (por quincena; Q1 intacto) + coherencia leaf.
- Procesadores: path Q2 por periodo con fail-fast que nombra ASE + reporte.
- Golden Capa A Q2 (canonico T0) + regresion Q1 verde + test Optimizado obligatorio.
- Resumen AjustesSfT por ASE en log/Serilog (delta minimo UI).

### 1.3 Out of Scope

Todo §0.2 (2.6, 2.7, INTERVENTORIA, ANT EXT-REV, insert/delete de filas, DetRetri). Ademas: motor Excel/COM en CI; relectura de TXT diarios de EFC (no existen en insumos); restyle UI; framework DI.

### 1.4 Resultado de negocio

El Ingeniero ejecuta el modo 5 ASE sobre `REMUNERACION 2026072` con la plantilla canonica; obtiene `Remuneracion 202607-2 Total.xlsx` con D85:D89 (AJUSTES-SF-T) correctos post-Excel y D104:D109 actualizados con ajustes; Q1 sigue produciendo `AjustesSfT = 0` bit-a-bit igual que antes; el log audita saldos-nota y retribucion por ASE; las pruebas demuestran `AjustesSfT` = fuente ±0.5 por ASE.

### 1.5 Base documental (origen funcional — citas por seccion)

| Documento | Seccion / instruccion | Que aporta a 2.5 |
|---|---|---|
| `Detalle de plantilla.docx` | Inst. hoja `AJUSTES - SF-T` (rango por ASE, composicion SALDOS POR NOTA + RETRIBUCION NEGATIVA), Inst. `SALDOS POR NOTA`, Inst. `RETRIBUCION NEGATIVA`, fila R5 → CONSOLIDADO D85:D89 | Layout de la cadena y semantica de cada tramo; base del mapa T0 |
| `Prompt Maestro Vo.docx` | Paso de pegado en valores de saldos por nota y de retribucion negativa + destino AJUSTES-SF-T/CONSOLIDADO | Que tramos se pegan en valores (celdas editables candidatas) vs calculados |
| Rector Propuesta | §9 it. 2.5 (alcance), §6 (preservacion de formulas, lectura directa, trazabilidad), §10 CA + tolerancia ±0.5 | Rector normativo |
| `Proceso de Recaudo.docx` | — | **NO aplica** a esta HU (flujo de recaudo por banco/canal, ya cubierto en HU-09) |

---

## 2. DESIGN

### 2.1 Tablas verificadas (contratables desde el dia uno)

**Contratos que NO cambian (V1/V7):**

| Elemento | Estado | Tratamiento 2.5 |
|---|---|---|
| `ConsolidadoAse.TotalAse` (getter con `AjustesSfT`) | Existe y ya suma ajustes | Sin cambio; `TotalAse`/`GranTotal` Q2 se actualizan solos al poblar `AjustesSfT` |
| `CalcularConsolidado(periodo, datos-sin-ajustes)` fail-fast Q2 | Existe | **Intacto**: sigue lanzando sin ajustes (garantia contra calculo silencioso incompleto) |
| `ValidadorBasico`: `AjustesSfT == 0` Q1 + `GranTotal == Σ TotalAse` | Existe (3 sobrecargas) | Intacto; Q2 suma reglas por quincena |
| `BuscarArchivo` prefix-based sin fechas | Existe | Reutilizado por los finders nuevos |
| `BuscarBalance` base → Optimizado | Existe (HU-10) | Reutilizado; su cobertura Q2 pasa a obligatoria (V4) |
| Q1 golden `D85:D89 = 0` + `'AJUSTES - SF-T'!D47..D51` (valor 0 en Q1) | Evidencia HU-07T0 §0.5 | Baseline de regresion: Q1 no se mueve |

**Insumos Q2 verificados en disco (V3/V4/V5):**

| ASE | SaldosNotas (rango) | RetribucionNegativa (rango) | Balance |
|---|---|---|---|
| 1-Promoambiental | 01072026–31072026 | 01072026–31072026 | plano |
| 2-Lime | 01072026–31072026 | 01072026–31072026 | plano |
| 3-Ciudad Limpia | 01072026–31072026 | 01072026–31072026 | plano |
| 4-Bogota Limpia | 01072026–31072026 | **16072026–31072026** | plano |
| 5-Area Limpia | 01072026–31072026 | 01072026–31072026 | **solo Optimizado** |

> Tablas golden Q2 (D85:D89, `TotalAse`, `GranTotal` por ASE), composicion exacta D85:D89, layouts de headers de ambas fuentes y tablas de operandos las extrae T0 del cache golden + fuentes; este plan no las inventa.

### 2.2 Decisiones de arquitectura

| ID | Opcion elegida | Descartada | Por que |
|---|---|---|---|
| D1 | Overloads `Calcular(ase, r1, r2, r4, decimal ajustesSfT)` + `CalcularConsolidado(periodo, datos-con-ajustes)`; overloads viejos intactos (Q1 = 0, Q2-viejo sigue fail-fast) | Parametro opcional o flag `incluirAjustes` en la firma existente | OCP: el path certificado Q1 no cambia de firma ni de semantica; el fail-fast viejo queda como red de seguridad contra llamadas sin ajustes |
| D2 | AJUSTES-SF-T con **doble desenlace T0**: (a) si hay operandos editables demostrados → se escriben en la misma pasada; (b) si todo es formula → visibles al mapa protegido y HU-11 no agrega escrituras | Presuponer celdas escribibles antes de T0 | Doctrina BCE-H (HU-10 D2): H era formula → protegida sin rebase. Ambos desenlaces usan los mismos gates; solo cambia el mapa escribible |
| D3 | Modelos `SaldosNotasAseInputs` / `RetribucionNegativaAseInputs` con `TieneColumnaEspeciales` + `ServEspK` (patron R2 literal) y mapeo por titulo de columna | Columnas fijas o asumir "Especiales" siempre presente | V6 + patron R2 (`SaldosFavorR2.TieneColumnaEspeciales`; ASE4 sin Especiales): la ausencia es legitima y vale 0, no fallo |
| D4 | Locator: `BuscarSaldosNotas` (prefijo `SaldosaFavorAplicadosPorNotas`) + `BuscarRetribucionNegativa` con **match normalizado sin diacriticos** (stem `retribucionnegativa`); ambos sin fechas | Un prefijo con ` cruzó` literal o con rango de fechas | V3 (rango ASE4 distinto) + V9 (encoding degradado en disco): fecha o acento literal = bug garantizado en ASE4 |
| D5 | Validador Q2 por quincena: si `Periodo.NumeroQuincena == 1` → reglas viejas (`AjustesSfT == 0`); si `== 2` → gate `AjustesSfT == TotalAjustes-fuente ±0.5` por ASE (matcheo estricto `Single` por Id) | Unificar reglas sin distinguir quincena | Q1 es regresion blindada (V7); Q2 es capacidad nueva. Mezclarlas romperia 78/78 |
| D6 | Misma pasada de escritura HU-07..HU-10 (overload lista; un `File.Copy` + validacion pre/post con mapa ampliado) | Segundo pase / writer separado | La superficie nueva son celdas del mismo modo leaf (o cero celdas, D2b). Un segundo pase romperia la atomicidad certificada y el hash A4 |
| D7 | `WorkbookLeafCellMapAjustesSfT`: dict explicito por `Ase.Id` (editables T0 + visibles esperados + composicion D85:D89 congelada), hermano de los mapas HU-07..HU-10 | Offsets aritmeticos o reutilizar el mapa R2 | Estructura multi-ASE apilada verticalmente (prohibidos los offsets); la cadena AJUSTES tiene su propia aritmetica dictada por T0 |
| D8 | Golden Q2 = **un canonico** (T0-0.1) + Capa A (§2.7); la segunda plantilla = control/estructura, nunca oraculo | Promediar o "elegir en el test" | Dos oraculos = merge verde en falso. T0 fija uno + SHA256 |

### 2.3 Escritura 2.5 (misma sesion atomica)

```text
File.Copy plantilla (canonica T0-0.1) → salida (una vez, igual que HU-07..HU-10)
  └─► ValidarFormulasProtegidas (mapa HU-07..HU-10 + mapa 2.5: AJUSTES-SF-T visibles D47..D51,
      D85:D89, INTERVENTORIA L25:N31, ANT EXT-REV, DetRetri/DetValiRetri Q2, resto 2.6–2.7)
        └─► EscribirCeldasLeaf HU-07..HU-10 (intactas)
        └─► EscribirCeldasAjustesSfT 2.5 (solo editables T0; en desenlace D2(b): nada)
              └─► revalidar formulas protegidas → guardar
```

Ante cualquier fallo: borrar salida parcial (patron existente). Plantilla origen jamas mutada (hash A4 se mantiene y se extiende al mapa 2.5). Sin capacidad de filas: fail-fast honesto (HU-07), nunca insert/delete.

### 2.4 Dominio (Core, sin deps)

```csharp
public sealed class SaldosNotasAseInputs  // un ASE: SALDOS POR NOTA
{
    public Ase Ase { get; set; } = new();
    public bool TieneColumnaEspeciales { get; set; }  // patron R2; ASE4 = false
    public decimal ServEspK { get; set; }             // 0 si ausente
    public IReadOnlyDictionary<string, decimal> Celdas { get; set; } = ...; // refs T0 → valor
    // TotalSaldosNotas: propiedad calculada (aritmetica T0; analoga a TotalOportuno R2)
}

public sealed class RetribucionNegativaAseInputs  // un ASE: RETRIBUCION NEGATIVA
{
    public Ase Ase { get; set; } = new();
    public IReadOnlyDictionary<string, decimal> Celdas { get; set; } = ...; // refs T0 → valor (negativos por componente)
    // TotalRetribucionNegativa: propiedad calculada (negativa; aritmetica T0)
}

public sealed class AjustesSfTInputs  // un ASE: composicion para D85:D89
{
    public Ase Ase { get; set; } = new();
    public SaldosNotasAseInputs SaldosNotas { get; set; } = new();
    public RetribucionNegativaAseInputs RetribucionNegativa { get; set; } = new();
    // TotalAjustes: propiedad calculada con la composicion congelada por T0-0.3
}
```

`WorkbookLeafInputs` suma `AjustesSfT: AjustesSfTInputs?` (`null` = comportamiento HU-10 intacto / Q1 puro; compatibilidad hacia atras por construccion). `ICalculoRemuneracion`: overloads D1 (contrato intacto, implementacion extendida). `ProcesadorPeriodo`: pasos 2.5 integrados al flujo por ASE en modo Q2 (leer saldos-nota + retribucion → leaf → validar gates D5) con fail-fast que nombra ASE **y reporte**; una escritura al final (G8). Modo Q1: pasos 2.5 se omiten (`AjustesSfT = null`, mismos asserts que hoy).

### 2.5 `IValidador` 2.5 (sin reabrir HU-04..HU-10)

1. Todo lo HU-07..HU-10 intacto (matcheo estricto por `Ase.Id`, `GranTotal = Σ`, gates leaf/empresa/banco/balance).
2. Nuevo, **solo si `Periodo.NumeroQuincena == 2`**: por cada ASE, `consolidado.AjustesSfT == leaf.AjustesSfT.TotalAjustes` ±0.5 (matcheo `Single` por Id; si `AjustesSfT` es null → error que nombra el ASE).
3. Nuevo Q2: `TotalAse` aritmetico por ASE ya incluye ajustes via getter (re-verificar con test, sin cambio de codigo salvo bug).
4. Q1 sin cambios: `AjustesSfT == 0` en los 5 + mensaje actual (regresion 78/78 ciega a 2.5).
5. El validador NO abre `.xlsx` (igual que HU-06..HU-10). La composicion D85:D89 va a Capa A (dominio vs cache golden), no a gates.
6. Coherencia `WorkbookLeafCoherence`: `ValidarContraResultado` Q2 itera ajustes con matcheo estricto (misma regla HU-07, extendida).

### 2.6 UI — Visual Design Intent (delta minimo)

Densidad Balanced, mismos GroupBoxes, sin restyle/colores/iconos. El `txtLog` agrega, por cada ASE en modo Q2, una linea AJUSTES (saldos-nota / retribucion-negativa / total, "esperado post-Excel") + D85:D89 esperado. Serilog: mismos eventos HU-07..HU-10 con propiedad `Hoja = "AJUSTES - SF-T"`. Sin nuevos controles (el periodo Q2 ya se selecciona; `NombreArchivo` ya resuelve `Remuneracion 202607-2 Total.xlsx`).

### 2.7 Golden Capa A extendida a 2.5 (honestidad HU-06..HU-10)

| # | Que | Contra que | Tol |
|---|---|---|---|
| A1 | Celdas 2.5 escritas en la **salida** (solo editables T0; en D2(b): vacio — el test lo declara) | Mismas celdas **leaf** del golden canonico | ±0.5 |
| A2 | `TotalAjustes` de **dominio** por ASE (composicion T0-0.3) | Cache golden D85:D89 + `AJUSTES - SF-T` D47..D51 | ±0.5 |
| A3 | `AJUSTES - SF-T` visibles + D85:D89 + INTERVENTORIA L25:N31 + ANT EXT-REV + DetRetri/DetValiRetri siguen siendo formula en la salida | Estructura | n/a |
| A4 | SHA256 plantilla canonica igual antes/despues | — | n/a |
| A5 | **Prohibido** comparar cache de formula de la salida vs golden; **prohibido** usar `TotOpt` HU-02 o totales Q1 como oraculo de ajustes | — | prohibido |
| A6 | `TotalAse`/`GranTotal` de dominio (con ajustes) vs cache golden D104:D109 | Cache golden canonico | ±0.5 |
| A7 (nuevo) | Q1 intacto: golden Q1 D85:D89 = 0 + `AjustesSfT = 0` re-assert en los tests Q1 existentes (sin duplicar suite) | Golden Q1 | ±0.5 |

Capa B (manual Excel: abrir, recalcular, comparar D85:D89 + D104:D109 vs golden canonico) fuera de CI, protocolo §5.3. **Capa B manual queda como accion del usuario** (disciplina del proyecto).

### 2.8 File changes previstos

| Archivo | Accion | Motivo |
|---|---|---|
| `Remuneracion.Core/Models/SaldosNotasAseInputs.cs` | Crear | SALDOS POR NOTA por ASE (patron R2) |
| `Remuneracion.Core/Models/RetribucionNegativaAseInputs.cs` | Crear | RETRIBUCION NEGATIVA por ASE (negativos por componente) |
| `Remuneracion.Core/Models/AjustesSfTInputs.cs` | Crear | Composicion por ASE (aritmetica T0-0.3) |
| `Remuneracion.Core/Models/WorkbookLeafInputs.cs` | Modificar | Sumar `AjustesSfT` (nullable, default null) |
| `Remuneracion.Core/Interfaces/ICalculoRemuneracion.cs` | Modificar | Overloads con ajustes (D1); firmas viejas intactas |
| `Remuneracion.Core/Services/CalculoRemuneracion.cs` | Modificar | Implementar overloads Q2; fail-fast viejo intacto |
| `Remuneracion.Core/Interfaces/IWorkbookLeafInputReader.cs` | Modificar | `LeerSaldosNotas` + `LeerRetribucionNegativa` (overloads) |
| `Remuneracion.Core/Interfaces/IValidador.cs` | Modificar | Overload/gates Q2 por quincena (§2.5) |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Modificar | Gates D5 (Q2 exige ajustes; Q1 exige 0 — intacto) |
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` | Modificar | Path Q2 por ASE (leer 2.5 → leaf → validar), fail-fast ASE+reporte |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapAjustesSfT.cs` | Crear | Mapa explicito por `Ase.Id` (congelado T0, incluye composicion D85:D89) |
| `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` | Modificar | Lectura header-driven 2.5 (Especiales opcional, mapeo por titulo) |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs` | Modificar | Coherencia ajustes-vs-consolidado Q2, matcheo estricto |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Escribir editables 2.5 en la misma pasada + mapa protegido ampliado |
| `Remuneracion.Infrastructure/FileSystem/ArchivoFuenteLocator.cs` | Modificar | `BuscarSaldosNotas` + `BuscarRetribucionNegativa` (agnosticos a rango/diacriticos, D4) |
| `Remuneracion.Core/Interfaces/ILocalizadorArchivosAse.cs` | Modificar | Firmas de los dos finders nuevos |
| `Remuneracion.WinForms/Form1.cs` | Modificar | Resumen AJUSTES por ASE en log (§2.6) |
| `Remuneracion.IntegrationTests/GoldenAjustesSfTQ2Tests.cs` | Crear | Capa A Q2 (§2.7, golden canonico T0-0.1) |
| `Remuneracion.IntegrationTests/AjustesSfTTests.cs` | Crear | Dominio: gates Q2, Especiales ausente = 0, composicion T0, mismatch nombra ASE+reporte |
| `Remuneracion.IntegrationTests/ProcesadorPeriodoTests.cs` | Modificar | Casos Q2 (integracion periodo + negativas rango ASE4/diacriticos) |
| `Remuneracion.IntegrationTests/Insumos.cs` | Modificar | Helpers `SaldosNotas(aseId)` + `RetribucionNegativa(aseId)` + `PeriodoQ2()` + golden canonico (T0-0.1) |

**No tocar (salvo bug blocker):** `IPlantillaWriter`/validation-only (HU-04); `IRecaudoReader` agregados HU-02; coherencia `F25`-Extemp HU-05; semantica single-ASE y multi-ASE HU-07; mapas HU-08/HU-09/HU-10; overloads viejos de `ICalculoRemuneracion`; reglas Q1 de `ValidadorBasico`; `requirements/` legado.

---

## 3. SPEC / Acceptance Criteria (trazabilidad Propuesta §10 + base docs)

### Requirement 1 — Desbloqueo Q2 con regresion Q1 (CA-2; HU-03 fail-fast)

El sistema **MUST** calcular `AjustesSfT` real en Q2 via los overloads nuevos y **MUST NOT** alterar el path Q1: overload viejo con Q2 sigue lanzando `CalculoInvalidoException`; Q1 sigue con `AjustesSfT = 0`.

- GIVEN tuplas Q2 **con** ajustes → WHEN `CalcularConsolidado(periodoQ2, datosConAjustes)` → THEN 5 consolidados con `AjustesSfT` = fuente ±0.5 y `GranTotal = Σ TotalAse`.
- GIVEN periodo Q2 con overload **viejo** (sin ajustes) → THEN `CalculoInvalidoException` (red de seguridad intacta).
- GIVEN periodo Q1 (cualquiera de los 78 tests existentes) → THEN comportamiento bit-a-bit identico (`AjustesSfT = 0`, D85:D89 = 0).

### Requirement 2 — Lectura header-driven con Especiales opcional (CA-1; Detalle de plantilla Inst. SALDOS POR NOTA / RETRIBUCION NEGATIVA)

El sistema **MUST** leer ambas fuentes por headers/titulos (búsqueda dinamica, nunca fila/columna fija) con columna "Especiales" opcional (ausente → 0, patron R2; ASE4-Bogota Limpia). Los valores de Retribucion son **negativos por componente**. **MUST NOT** inventar valor si falta un header esperado → fallo que nombra ASE + reporte.

- GIVEN fuente ASE4 sin "Especiales" → THEN `TieneColumnaEspeciales = false`, `ServEspK = 0`, resto leido normal.
- GIVEN Retribucion ASE1 → THEN totales negativos por componente ±0.5 vs fuente.

### Requirement 3 — AJUSTES-SF-T formulada validada, no escrita (CA-3/CA-4; doctrina BCE-H HU-10)

El sistema **MUST** tratar los visibles `AJUSTES - SF-T` D47..D51 como formulas protegidas (nunca objetivo de escritura) y **MUST** validar su composicion exacta segun T0-0.3. La composicion D85:D89 **MUST** venir del T0, nunca de hipotesis.

- GIVEN mapa T0 con desenlace D2(a) → THEN solo cambian editables evidenciados; A3 verde.
- GIVEN desenlace D2(b) (todo formula) → THEN cero escrituras nuevas y la HU igual cierra (calculo + validacion + golden).

### Requirement 4 — Golden Q2 canonico + D85:D89 y totales (CA-3/CA-5)

Las pruebas **MUST** ejecutar la matriz §2.7 contra el golden canonico fijado por T0-0.1. **MUST NOT** comparar cache de formula de la salida vs golden (A5); **MUST NOT** usar la segunda plantilla como oraculo.

- GIVEN salida Q2 generada → THEN A1 (leafs) + A2 (D85:D89 dominio vs cache) + A6 (totales) ±0.5.
- GIVEN la plantilla de trabajo es el propio golden → THEN ningun assert compara cache de formula de la salida.

### Requirement 5 — Balance-Optimizado obligatorio (CA-1; V4/V8)

El path Q2 **MUST** resolver ASE5 via `BuscarBalance` fallback Optimizado (ya existe) con test dedicado **obligatorio** (no opcional): `BalanceOptimizado_Q2_ResuelvePorSegundoPrefijo`. Si no existe ni base ni Optimizado → fail-fast que nombra el ASE.

### Requirement 6 — Locator agnostico a rango y diacriticos (CA-1; V2/V3/V9)

Los finders nuevos **MUST** resolver los 5 ASE pese al rango ASE4 (1607–3107) y a los diacriticos del nombre (`RetribucionNegativa`). **MUST NOT** codificar fechas ni literales con acento fragil.

- GIVEN carpeta ASE4 (Retribucion rango 1607–3107) → THEN resuelve por prefijo sin fallo.
- GIVEN nombre con `o` acentuada en disco → THEN match normalizado resuelve igual.

### Requirement 7 — Guardas fuera de alcance (CA-4)

`INTERVENTORIA` L25:N31, `ANT EXT-REV`, DetRetri/DetValiRetri Q2 y resto 2.6–2.7 **MUST** estar en el mapa protegido (fallan la escritura si alguna deja de ser formula). Si la plantilla no alcanza para el detalle Q2 → fail-fast honesto (HU-07), **MUST NOT** insert/delete de filas.

| CA §10 | HU-11 |
|---|---|
| CA-1 | Lee saldos-nota + retribucion de los 5 ASE Q2 (fail-fast nombra ASE+reporte; locator agnostico) |
| CA-2 | `AjustesSfT` = fuente ±0.5 por ASE (agregados HU-02 ≠ valores de ajuste) |
| CA-3 | Capa A Q2 (D85:D89 + totales); AJUSTES-SF-T correcta post-Excel (Capa B manual residual) |
| CA-4 | Reassert mapa ampliado + plantilla canonica no mutada (hash) |
| CA-5 | Gates Q2 por quincena + coherencia leaf ajustes |
| CA-6 | Serilog + log AJUSTES por ASE |
| CA-7 | Mismo flujo 5-ASE + resumen AJUSTES (Q1 sin cambios visibles) |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 700–1100 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 T0 + golden canonico + mapa → PR2 dominio + calculo → PR3 readers + locator → PR4 writer + validador + periodo/UI → PR5 golden Q2 + Optimizado + regresion |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes. Chained PRs recommended: Yes. Chain strategy: pending.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 0 | Discovery T0 + golden canonico + mapa congelado | PR 1 | Bloquea todo; solo lectura + datos |
| 1 | Modelos + overloads `ICalculoRemuneracion` Q2 | PR 2 | Depende de PR 1 (composicion T0-0.3) |
| 2 | Readers 2.5 + locator agnostico | PR 3 | Depende de PR 1 |
| 3 | Writer protegido + validador Q2 + periodo + UI | PR 4 | Depende de PR 1 |
| 4 | Golden Q2 + Optimizado obligatorio + regresion Q1 | PR 5 | Depende de PR 2–4; 78 tests previos verdes |

### Phase 0 — Discovery T0 (BLOQUEANTE, solo lectura, sin cambiar escritura)

- [ ] 0.1 Fijar el **golden canonico Q2**: comparar las dos plantillas 202607-2 (n° hojas = 40?, hoja `AJUSTES - SF-T` presente en ambas?, D85:D89 formulas?, cual matchea `Remuneracion 202607-2 Total.xlsx` como referencia). Registrar eleccion + SHA256; la otra queda como control.
- [ ] 0.2 Clasificar valor-vs-formula celda por celda de la cadena AJUSTES-SF-T en el canonico: hoja `AJUSTES - SF-T` (visibles D47..D51 por ASE + operandos SALDOS POR NOTA y RETRIBUCION NEGATIVA), INTERVENTORIA L25:N31, ANT EXT-REV. Registrar `<f>` y `<v>` de cada una + veredicto D2(a)/D2(b).
- [ ] 0.3 Congelar la **composicion exacta D85:D89** (formula de cada D85..D89 + aritmetica de dominio equivalente para A2). **Nada de D85:D89 entra al codigo sin esta tabla.**
- [ ] 0.4 Verificar a nivel bytes los nombres reales en disco (`SaldosaFavorAplicadosPorNotas_*`, `RetribucionNegativa_*` con diacriticos) y fijar la normalizacion del matcher (lowercase + strip-diacritics + prefijo sin fechas); verificar que el prefijo resuelve ASE4 (rango 1607–3107) y los 5/5.
- [ ] 0.5 Dumpear headers de ambas fuentes por ASE (columna "Especiales" presente/ausente — confirmar ASE4 ausente; columna de negativos por componente en Retribucion; etiquetas de totales: ¿"Total General"? ¿"Grand Total"?). Fijar mapeo por titulo + regla "ausente Especiales = 0, ausente resto = fallo".
- [ ] 0.6 Verificar ASE5-Q2 Balance solo-Optimizado resuelve por `BuscarBalance` fallback (hacer obligatorio el test) + paridad de layouts Q2-vs-Q1 en R1/R2/R4/Banco/Balance (si alguno diverge, recorte escrito, no invencion).
- [ ] 0.7 Congelar `WorkbookLeafCellMapAjustesSfT` (editables + visibles + protegidas incl. INTERVENTORIA/ANT EXT-REV/DetRetri Q2) + check de capacidad de filas (si no alcanza: fail-fast honesto, sin insert/delete). **Nada entra al codigo sin esta tabla.**
- [ ] 0.8 Confirmar baseline Q1: D85:D89 = 0 y `AjustesSfT = 0` en golden Q1 (regresion ciega).

### Phase 1 — Dominio (modelos + calculo)

- [ ] 1.1 `SaldosNotasAseInputs.cs` + `RetribucionNegativaAseInputs.cs` + `AjustesSfTInputs.cs` (§2.4; totales como propiedades calculadas con aritmetica T0-0.3) + extender `WorkbookLeafInputs` (nullable, default null).
- [ ] 1.2 Overloads `ICalculoRemuneracion` (D1) + implementacion Q2 en `CalculoRemuneracion` (Q2 exige tupla con ajustes; Q1-viejo y fail-fast-viejo intactos).
- [ ] 1.3 Tests in-memory de calculo: Q2 con ajustes cuadra `TotalAse`/`GranTotal`; Q2-viejo lanza; Q1 identico.

### Phase 2 — Lectura (readers + locator)

- [ ] 2.1 `ILocalizadorArchivosAse`: `BuscarSaldosNotas` + `BuscarRetribucionNegativa` (D4: sin fechas, match sin diacriticos) + implementacion en `ArchivoFuenteLocator`.
- [ ] 2.2 `IWorkbookLeafInputReader`: `LeerSaldosNotas` + `LeerRetribucionNegativa` (header-driven, Especiales opcional, negativos por componente; fallo nombra ASE+reporte).
- [ ] 2.3 `Insumos.cs`: helpers `SaldosNotas(aseId)`, `RetribucionNegativa(aseId)`, `PeriodoQ2()`, golden canonico (T0-0.1).

### Phase 3 — Orquestacion + UI delta minimo

- [ ] 3.1 `ProcesadorPeriodo`: path Q2 (resolver 5 carpetas Q2 + 2 prefijos nuevos por ASE; leer R1/R2/R4 + 2.5 + resto de hojas; `CalcularConsolidado` overload nuevo; `Validar` Q2; una escritura). Q1 intacto.
- [ ] 3.2 `ValidadorBasico` + `WorkbookLeafCoherence`: gates §2.5 (Q2 exige ajustes + matcheo estricto; Q1 exige 0 — intacto).
- [ ] 3.3 `Form1`: resumen AJUSTES por ASE en log (§2.6) + Serilog `Hoja = "AJUSTES - SF-T"`.
- [ ] 3.4 `OpenXmlPlantillaWriter`: mapa protegido ampliado + escritura de editables T0 (o cero celdas en D2b) en la misma pasada.

### Phase 4 — Pruebas y evidencia

- [ ] 4.1 `AjustesSfTTests` (in-memory + fuentes Q2 reales sin salida): gate Q2 ok, Especiales ausente = 0, header ausente = fallo ASE+reporte, composicion T0 exacta, mismatch nombra ASE, TotOpt-vs-ajuste prohibido (test que intenta la confusion y falla).
- [ ] 4.2 `GoldenAjustesSfTQ2Tests`: matriz §2.7 (5 ASE, fixtures Q2, golden canonico; segunda plantilla nunca oraculo).
- [ ] 4.3 **OBLIGATORIO** `BalanceOptimizado_Q2_ResuelvePorSegundoPrefijo` + negativa (sin base ni Optimizado → fail-fast nombra ASE).
- [ ] 4.4 Negativas locator: ASE4 rango 1607–3107 resuelve; diacriticos resuelven; salida == plantilla (no in-place); overwrite cancelado.
- [ ] 4.5 Los 78 tests + harness 24/24 existentes verdes; build 0 warnings; CRLF; sin commit.

### Phase 5 — Documental

- [ ] 5.1 Capa B manual §5.3 ejecutada una vez sobre el canonico y evidenciada (accion del usuario; sin fingirla como gate de merge).
- [ ] 5.2 Cierre deja explicito el frente 2.6 (siguiente HU propuesta: DetRetri Q2 / escritura DetRetri).

---

## 5. Test Strategy

| Capa | Que | Como |
|---|---|---|
| Unidad | Gates Q2 (ajustes = fuente, Q1 = 0 intacto, `Single` por Id) | In-memory, composicion T0-0.3 |
| Unidad | Calculo overloads (Q2 abre, Q2-viejo lanza, Q1 identico) | In-memory |
| Unidad | Readers 2.5 (headers dinamicos, Especiales opcional, negativos) | Fuentes Q2 reales, sin Excel de salida |
| Integracion | Periodo Q2 (5 carpetas reales, salida temp) | Insumos Q2, fail-fast ASE+reporte |
| Golden Capa A Q2 | Matriz §2.7 | OpenXML read-only + aritmetica dominio; canonico T0-0.1 (A5 aplica) |
| Regresion | 78/78 + harness 24/24 Q1 verdes | Suite existente, sin cambios |
| UI | Resumen AJUSTES | Funcional manual (sin harness) |
| Capa B | D85:D89 + D104:D109 post-Excel | Manual — §5.3 |

### 5.1 Fixtures

- Golden/plantilla Q2: canonico fijado por T0-0.1 (una de las dos plantillas 202607-2; la otra = control).
- Fuentes Q2: `Docs/Insumos/REMUNERACION 2026072/{1..5}-*/` (`SaldosaFavorAplicadosPorNotas_*`, `RetribucionNegativa_*` + resto de reportes para paridad T0-0.6; ASE5 Balance solo-Optimizado).
- Regresion: golden + fuentes Q1 (intactos).
- Referencia: composicion T0-0.3 + tablas T0-0.5, tolerancia ±0.5.

### 5.2 Casos negativos obligatorios (nombran ASE y reporte)

Header "Especiales" ausente en ASE4 (vale 0, no falla); header clave ausente en ASE2 (falla ASE+reporte); Retribucion con rango ASE4 1607–3107 (resuelve); diacriticos en disco (resuelven); Q2 por overload viejo (lanza); Q1 con ajustes ≠ 0 (bloquea, intacto); segunda plantilla Q2 como oraculo (prohibido — el test lo demuestra fallando); `TotOpt` HU-02 como valor de ajuste (prohibido); salida == plantilla (no in-place); plantilla sin filas para detalle Q2 (fail-fast honesto, sin insert/delete).

### 5.3 Protocolo manual Capa B (no CI — accion del usuario)

1. Generar salida Q2 a ruta distinta del canonico. 2. Abrir en Excel, recalcular. 3. Comparar D85:D89 + D104:D108 + D109 vs canonico ±0.5. 4. Evidencia adjunta. No es gate de merge.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | Ajustes = extension de datos (mapa + inputs + readers + gates); calculo/periodo/writer conservan su rol HU-07..HU-10. |
| **O** | Se agregan modelos/mapas/overloads; el path Q1 funciona con `AjustesSfT = null` (abierto sin modificar). El fail-fast viejo queda como red. |
| **L** | `CalculoRemuneracion` suma overloads sin cambiar el comportamiento Q1 ni el fail-fast del overload viejo. |
| **I** | `AjustesSfTInputs` separado de `BalanceScInputs`/`ReporteBancoInputs`; reader/validador/locator crecen por overload. |
| **D** | Core define ajustes/inputs/gates; Infrastructure/WinForms componen. Sin nuevas deps. |

### 6.2 Best Practices

- La verdad del workbook manda: AJUSTES-SF-T formulada ⇒ cero escritura directa (Rector §6 preservacion; doctrina BCE-H HU-10).
- La composicion D85:D89 se prueba contra el golden, no se asume (honestidad §0.1).
- Mapa explicito por `Ase.Id` verificado, no offsets ni filas/columnas fijas (Rector §11.3).
- Locator agnostico a rango y diacriticos (V3/V9): fechas y acentos literales prohibidos.
- Una escritura atomica; plantilla canonica nunca mutada; hash A4 extendido.
- Golden honesto Q2 (OpenXML no recalcula; A5/A6/A7).
- Fail-fast nombra ASE+reporte; sin salida certificada ante fallo; sin insert/delete de filas.
- Q1 blindada por construccion (overloads viejos + 78/78 como red).

### 6.3 Performance

- 5 ASE × 2 lecturas header-driven + ~10 escrituras (o cero en D2b) en la misma sesion OpenXML + 5 resoluciones por prefijo. Irrelevante a esta escala; `Task.Run` existente para no congelar el form.

**Veredicto:** APROBADO como it. 2.5 de Fase 2 **si** T0 congela el mapa con evidencia (incluidas composicion D85:D89 y golden canonico) y se acepta CA-3 parcial (Capa A en CI, Capa B manual como accion del usuario).

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigacion |
|---|---|---|
| La composicion D85:D89 no cierra ±0.5 contra ninguna hipotesis (cadena no entendida) | Media | T0-0.3 declara veredicto fallido; HU recortada y escalada al Ingeniero; prohibido inventar composicion |
| Las dos plantillas 202607-2 difieren en layout AJUSTES y el canonico es ambiguo | Media | T0-0.1 con criterios escritos (40 hojas, D85:D89, match con referencia); si ambas valen, se fija una y se documenta la otra como control |
| Algun operando AJUSTES es formula no editable (caso inverso a D2a) | Media | D2 doble desenlace: el gate es identico; solo cambia el mapa escribible. T0-0.2 lo decide antes de codificar |
| Layouts Q2-vs-Q1 divergen en R1/R2/R4/Banco/Balance | Media | T0-0.6; si el hueco es real, recorte a los reportes verificados con rebase, no invencion |
| Fuente 2.5 no trae desglose por componente/ASE esperado | Media | T0-0.5; el fallo nombra ASE+reporte (fail-fast), no se finge mapeo |
| Ceros legitimos confundidos con "falta de lectura" (Q2 puede traer ceros) | Alta | El reader distingue "leido 0" de "header ausente" (fallo); tests de ceros explicitos |
| Rango ASE4 (1607–3107) rompe un matcher con fechas | Alta | D4 + test negativo dedicado; prohibido codificar fechas |
| Diacriticos en disco rompen el prefijo literal | Alta | D4 (normalizacion) + T0-0.4 a nivel bytes + test dedicado |
| Inflar a 2.6–2.7 dentro de esta HU (DetRetri tienta, INTERVENTORIA tienta) | Media | §0.2 out-of-scope + protegidas, no implementadas; rechazar PRs que lo metan |
| Comparar cache de salida vs golden y "cerrar" CA-3 en falso (doble plantilla agrava) | Alta | A5 + canonico unico; este plan lo prohibe |
| Q1 regresion rota por un overload mal preservado | Media | 78/78 + 24/24 como red en PR5; Q1 intacto por construccion (D1/D5) |

---

## 8. Rollback

- Revertir PRs en orden inverso (golden → periodo/UI/validador → writer → readers/locator → dominio → mapa/T0).
- HU-04..HU-10 intactas sin esta HU: `AjustesSfT = null` = comportamiento HU-10 puro; Q2 sigue fail-fast honesto.
- `Docs/Insumos/` untracked: jamas como destino; pisada se restaura desde backup.
- Si T0 demuestra hueco de fuente o composicion fallida (Riesgos 1–2), recorte con rebase, no invencion.

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §9 it. 2.5 / §6 / §10 (±0.5) + Detalle de plantilla y Prompt Maestro como origen funcional (§1.5). 2.6–2.7, INTERVENTORIA, ANT EXT-REV y Q2-D detRetri quedan fuera. Proceso de Recaudo NO aplica.
2. Desbloqueo Q2 por overloads nuevos (D1); overloads viejos y fail-fast intactos como red; Q1 bit-a-bit identica.
3. AJUSTES-SF-T = hoja FORMULADA validada, no escrita (doctrina BCE-H HU-10); doble desenlace T0 sin rebase (D2).
4. Lectura header-driven con "Especiales" opcional (patron R2; ASE4 sin Especiales → 0); negativos por componente en Retribucion; mapeo por titulo.
5. Composicion exacta D85:D89 dictada por T0-0.3 contra el golden; nunca asumida.
6. Golden Q2 unico y canonico (T0-0.1); segunda plantilla = control, nunca oraculo.
7. Locator agnostico a rango (ASE4 1607–3107) y a diacriticos (normalizacion verificada a nivel bytes).
8. Test Balance-Optimizado OBLIGATORIO (ASE5-Q2 solo trae Optimizado).
9. Una sola escritura atomica; sin insert/delete de filas (fail-fast honesto si la plantilla no alcanza).
10. Golden Capa A Q2 con honestidad HU-06..HU-10; Capa B manual residual como accion del usuario.
11. UI delta minimo + Serilog AJUSTES por ASE; OPA = Ejecutar.
12. Apply espera aprobacion + PRs encadenados (Unidad 0–4, §4).

**Listo para aprobacion del Ingeniero. No implementar hasta OK.**
