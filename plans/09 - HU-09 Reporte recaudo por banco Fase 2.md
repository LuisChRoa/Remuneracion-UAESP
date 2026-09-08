# Plan 09 — HU-09: REPORTE RECAUDO x BANCO (Fase 2, it. 2.3)

> **Historia:** diligenciar la hoja `REPORTE RECAUDO x BANCO` para los 5 ASE desde el "Resumen Recaudo Aplicado Por Servicio" al final de cada `ReportePagosxBanco_*.xlsx`, dentro de la misma escritura atómica multi-ASE certificada por HU-07/HU-08.
> **Rector (irrenunciable):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 2 it. **2.3 REPORTE RECAUDO x BANCO**, §6 arquitectura (capas, preservación de fórmulas, Serilog, lectura directa de fuente), §10 CA + tolerancia ±0.5. **Entra SOLO 2.3.** Salen 2.4 BCE, 2.5 AJUSTES-SF-T, 2.6 DetRetri, 2.7 validaciones cruzadas — son HUs posteriores. Quincena 2 sigue bloqueada.
> **Origen funcional:** los tres documentos base — `Detalle de plantilla.docx` (Inst. hoja `REPORTE RECAUDO x BANCO`), `Proceso de Recaudo.docx` (§ reporte de pago por banco / recaudo aplicado), `Prompt Maestro Vo.docx` (paso de pegado en valores del reporte por banco) — citados por sección en §1.5/§3.
> **Continuidad:** HU-01..HU-08 cerradas o en cierre (ruta leaf multi-ASE atómica single-write, mapas explícitos por `Ase.Id`, validador estricto, UI modo 5-ASE, Golden Capa A, build 0/0, tests 33/33, harness 24/24). Este plan NO reabre su semántica. Tándem con HU-08: las filas 59–80 validan contra las hojas `Recaudo *` que implementa HU-08.
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-08

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. El orquestador ya verificó el golden Q1 y una fuente Q1 (Promoambiental) contra el template; lo que queda abierto es el carácter valor-vs-fórmula de las filas 1–7 (col J trae TRUE/TRUE) y la ubicación exacta del "Resumen" en las otras 4 fuentes — por eso el plan incluye discovery T0 obligatorio (§4 Fase 0). El apply espera aprobación explícita.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (dump OpenXML del golden `Docs/Insumos/Remuneracion 202607-1 Total.xlsx`, hoja `REPORTE RECAUDO x BANCO` + lectura fuente ASE1 Q1 real):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | Template hoja `REPORTE RECAUDO x BANCO`: filas 1–6 consolidadas; **fila 7 trae flags TRUE/TRUE en col J** → hay fórmulas vivas ahí; T0 clasifica | Dump raw golden (orquestador) | Filas 1–7: NO prescribir escritura sin T0. Si son fórmulas, se protegen |
| V2 | 5 bloques ASE × 10 filas (etiqueta ASE, etiqueta "Resumen…", header, filas 1/2/3/7, Total), filas 9–58 | Dump golden | Superficie escribible candidata: bloques 9–58 + C59 |
| V3 | Golden ASE1 verificado: ENEL `16194517449 / 50286069 / 0 / 0`, OCCIDENTE `523782724.81 / 12466967.19 / 4226000 / 0` | Caché golden | Oráculo Capa A para ASE1; resto de ASEs los extrae T0-0.5 |
| V4 | Fuente Promoambiental Q1 = una sola `Sheet1`, 408 `<row>`, filas diarias 1–~413 (bancos ENEL E01/E06/E07/…, OCCIDENTE H23 CNB / O23 OFC / T23 OFC), subtotales por canal ~400–408, **"Resumen Recaudo Aplicado Por Servicio" final r414–419** | Lectura fuente real | Estrategia de lectura: **búsqueda dinámica de etiqueta desde el final**, nunca fila fija (el conteo de filas diarias varía por ASE) |
| V5 | Resumen fuente ASE1 = template ASE1 **EXACTO**: ENEL `16194517449+50286069+0=16244803518`, OCCIDENTE `523782724.81+12466967.19+4226000=540475692`, Total `16785279210` | Aritmética fuente vs caché golden | Gate de coherencia central: bloque template = resumen fuente ±0.5 |
| V6 | Variante de etiqueta real: fuente r402 dice **"3-APLICADOS A FINANCIACIONES"** (sin NUEVAS) vs template "3-APLICADOS A FINANCIACIONES NUEVAS" | Lectura fuente | T0 debe matchear por prefijo/normalización, nunca string exacto hardcodeado |
| V7 | C59 = `1` con texto "OJO/// DILIGENCIAR" (selector de quincena 1 o 2) | Caché golden | C59 es celda de **valor** diligenciable; su semántica exacta (¿la leen fórmulas?) la fija T0 |
| V8 | Filas 59–80 = validación por empresa ("Recaudo ENEL - ASE n" con TRUE/FALSE + diferencia; ej. ASE2 ENEL FALSE `397250`); fila 80 "TOTAL RECAUDO" `58447133168` | Caché golden | Filas 59–80 + TOTAL son candidatas a **fórmulas protegidas** (comparan contra hojas `Recaudo *` de HU-08); T0 lo confirma |
| V9 | Semántica de negocio: la hoja resume TODO el recaudo aplicado de la quincena (oportuno + extemporáneo), **NO afectado por anulaciones/reversiones de la misma quincena**; E59:E80 = diferencias = anulado/reversado misma quincena | Documentos base (§1.5) | El lector 2.3 IGNORA reversiones; el validador NO exige igualdad con `Recaudo *` (la diferencia es esperada y se reporta, no se falla) |
| V10 | Mapa de columnas fuente por ASE verificado (orquestador): LIME + BOGOTÁ LIMPIA → ENEL, NUEVO ESQUEMA, OCCIDENTE; AREA LIMPIA → ENEL, ENERBIT, OCCIDENTE; PROMOAMBIENTAL + CIUDAD LIMPIA → ENEL, OCCIDENTE | Lectura fuentes | `ReporteBancoInputs` usa mapa explícito por `Ase.Id`, no columnas fijas |
| V11 | Títulos filas 1–7: ENEL (= facturación conjunta ENEL), CIUDAD LIMPIA + EAB (= antiguo "EAAB + Ciud Limp" particionado), NUEVO ESQUEMA (= EAAB Reciprocidad esquema actual), ENERBIT, OCCIDENTE (= facturación directa Banco Occidente) | Documentos base + golden | El consolidado 1–7 es Σ por empresa sobre los 5 bloques (aritmética de dominio si son valores; verificación si son fórmulas — T0 decide) |
| V12 | Contratos listos para extender: overload lista multi-ASE (`IWorkbookLeafWriter`), `WorkbookLeafInputs` con listas, `Insumos.cs` + `CarpetasAse.Prefijos` cubren las 5 carpetas; `ArchivoFuenteLocator` resuelve por prefijo `ReportePagosxBanco_` | Lectura de código | 2.3 = **extensión del patrón HU-07/HU-08**, no arquitectura nueva |
| V13 | Naming fuente confirmado en disco: `ReportePagosxBanco_to_date01072026ddMMyyyy_to_date16072026ddMMyyyy___2026716134349474.xlsx` en cada carpeta `1-Promoambiental … 5-Área Limpia` | `Get-ChildItem` Insumos | Sin nuevos tipos de archivo; el locator solo suma un prefijo |

**NO verificado (y por eso T0 es bloqueante, §4 Fase 0):**

1. Carácter valor-vs-fórmula celda por celda de filas 1–7 (especialmente col J fila 7 y columnas de totales), bloques 9–58, C59, filas 59–80 y fila 80 TOTAL.
2. Ubicación y etiqueta exacta del "Resumen Recaudo Aplicado Por Servicio" en las fuentes ASE2..ASE5 (¿misma etiqueta? ¿misma posición relativa al final? ¿variantes?).
3. Columnas fuente exactas por ASE (nombres de header de empresa en el Resumen: ¿"ENEL", "NUEVO ESQUEMA", "ENERBIT", "OCCIDENTE" literales? ¿orden?).
4. Filas 1/2/3/7 del Resumen por ASE: ¿los 4 conceptos existen siempre o hay filas ausentes cuando el valor es 0?
5. Si C59 es referenciada por alguna fórmula (¿las filas 59–80 o el TOTAL dependen de C59?).
6. Si alguna celda objetivo de 2.3 depende de un externalLink roto.
7. Tablas golden de bloques ASE2..ASE5 (las extrae T0-0.5 del caché; este plan no las inventa).

> **Regla de hierro del plan:** ninguna dirección de celda de filas 1–7 entra al código sin pasar por T0. Lo ya verificado arriba (V1–V13) sí es contratable desde el día uno.

### 0.2 Mapeo al Rector (in vs out)

**Entra porque §9 it. 2.3 / §6 / §10 lo piden ahora:**

| Rector | Qué cubre HU-09 |
|---|---|
| §9 it. 2.3 | `REPORTE RECAUDO x BANCO`: bloques por ASE 9–58 + consolidado 1–7 + C59 + validación 59–80 (protegida) |
| §7.1 bloque escritura | "ReportePagosxBanco → pega recaudo por banco" (valores, misma pasada atómica) |
| §10 CA-1/CA-2 | Leer el Resumen por banco de los 5 ASE + coherencia bloque = fuente y consolidado = Σ ASE |
| §10 CA-4/CA-5/CA-6/CA-7 | Fórmulas intactas (mapa ampliado 2.3), validación por ASE, Serilog por ASE, UI sin cambios salvo resumen |

**Sale porque §9 lo asigna a 2.4–2.7 / Fase 3 (HUs posteriores):**

- 2.4 `BCE SC POR FACT.` (Balance Subsidio/Contribuciones).
- 2.5 AJUSTES-SF-T (`SALDOS POR NOTA` + `RETRIBUCION NEGATIVA`; Q1 sigue con `AjustesSfT=0`; quincena 2 sigue bloqueada en `CalculoRemuneracion`).
- 2.6 `DetRetri2026071` / `DetValiRetri2026071` (enteros redondeados como objetivo de escritura).
- 2.7 validaciones cruzadas (`VALIDACION_TOTAL`, `VALIDACION_RECIP/ENEL/…`, `Valida -*`, y la propia lógica de filas 59–80 más allá de protegerlas). Se protegen sus fórmulas; su lógica no se implementa.
- `INTERVENTORIA`, `ANT EXT-REV`, `ANTICIPOS USUARIOS`, `Informe AFaseo Recaudo` (se ignora para el consolidado, docx Inst. 9).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **Lectura dinámica del Resumen desde el final** (búsqueda de etiqueta, match por prefijo normalizado). Prohibida fila fija: el conteo de filas diarias varía por ASE (V4/V6). |
| G2 | **Filas 59–80 + TOTAL como fórmulas protegidas** (hipótesis V8; T0 la confirma o corrige). Esta HU no implementa su lógica (es 2.7). |
| G3 | **Filas 1–7: T0 decide valor-vs-fórmula.** Si son valores → se escriben con Σ de dominio. Si son fórmulas → se protegen y solo se verifica Σ. El plan soporta ambos desenlaces sin rebase (ver D2). |
| G4 | **Una sola llamada de escritura** por proceso (mismo overload lista HU-07/HU-08). Sin segundo pase (ver D3). |
| G5 | **Mapa explícito por ASE** (`WorkbookLeafCellMapReporteBanco`: ASE × empresa → refs editables), no offsets. Columnas fuente heterogéneas por ASE (V10). |
| G6 | **La diferencia filas 59–80 es informativa, no fallo**: E59:E80 = anulado/reversado misma quincena (V9). El validador la reporta; no exige cero. |
| G7 | **UI sin cambios funcionales**: el modo 5-ASE ya existe; 2.3 solo agrega líneas de resumen por banco al log. Sin restyle. |
| G8 | **Golden Capa A extendida a 2.3** con honestidad HU-06/HU-07/HU-08 (leaf salida vs leaf golden; visibles de dominio vs caché golden; nunca caché de salida vs golden). |

---

## 1. PROPOSE

### 1.1 Intent

Diligenciar la hoja `REPORTE RECAUDO x BANCO` para los 5 ASE dentro de la misma escritura atómica del período — bloques por ASE (filas 9–58) en valores desde el Resumen de cada `ReportePagosxBanco`, consolidado 1–7 según veredicto T0, C59 con la quincena — de modo que las filas 59–80 calculen solas por fórmulas contra las hojas `Recaudo *` (HU-08) y la coherencia bloque = fuente quede demostrada contra el golden Q1.

### 1.2 In Scope

- Discovery T0 (§4 Fase 0) + `WorkbookLeafCellMapReporteBanco` congelado solo con evidencia.
- Lectura del Resumen por banco desde las 5 carpetas (mismo patrón de 6 xlsx por carpeta; el banco es uno de ellos — sin nuevos tipos de archivo).
- Extensión de `WorkbookLeafInputs` + escritura en la misma pasada HU-07/HU-08; validación pre/post con mapa ampliado (filas 59–80, TOTAL y — según T0 — filas 1–7 como protegidas).
- Gate bloque = fuente por ASE y gate consolidado = Σ ASE en `IValidador` (+ coherencia leaf).
- Golden Capa A extendida a 2.3 + tests con insumos Q1 reales.
- Resumen por banco en log/Serilog (delta mínimo UI).

### 1.3 Out of Scope

Todo §0.2 (2.4–2.7, Fase 3). Además: quincena 2 (sigue lanzando `CalculoInvalidoException`); reescritura del path HU-07/HU-08 (se extiende, no se duplica); filas 59–80 como lógica implementada (solo protegidas + diferencia reportada); restyle UI; DI framework.

### 1.4 Resultado de negocio

El Ingeniero ejecuta el modo 5 ASE como hoy; obtiene el mismo workbook más la hoja `REPORTE RECAUDO x BANCO` diligenciada (bloques por ASE en valores, consolidado coherente, C59 = quincena); las filas 59–80 muestran las diferencias vs `Recaudo *` post-Excel (anulados/reversiones, esperadas ≠ 0); el log audita por ASE **y por empresa de facturación**; las pruebas demuestran bloque = fuente y consolidado = Σ ASE.

### 1.5 Base documental (origen funcional — citas por sección)

| Documento | Sección / instrucción | Qué aporta a 2.3 |
|---|---|---|
| `Detalle de plantilla.docx` | Inst. hoja `REPORTE RECAUDO x BANCO` (filas 9–58 por ASE, conceptos 1-APLICADOS A FACTURACION / 2-SALDOS A FAVOR GENERADOS / 3-APLICADOS A FINANCIACIONES NUEVAS / 7-APLICADOS A RECIBOS SERVICIOS ESPECIALES Res. 27 incl. RCD; E59:E80 diferencias por anulados) | Layout de la hoja y semántica de cada fila; base del mapa T0 |
| `Proceso de Recaudo.docx` | § reporte de pago por banco (recaudo aplicado por servicio y empresa; bancos/canales por EFC) | Origen de los valores: el Resumen del reporte por banco resume el recaudo aplicado (oportuno + extemporáneo), sin reversiones misma quincena |
| `Prompt Maestro Vo.docx` | Paso de pegado en valores del reporte por banco + títulos del consolidado (ENEL / CIUDAD LIMPIA+EAB / NUEVO ESQUEMA / ENERBIT / OCCIDENTE) | La hoja se pega TODA en valores (celdas editables); títulos 1–7 y correspondencia con empresas |
| Rector Propuesta | §9 it. 2.3 (alcance), §6 (preservación de fórmulas, lectura directa, trazabilidad), §10 CA + Apéndice B (operadores por fuente) | Rector normativo; Apéndice B compatible con V10 |

---

## 2. DESIGN

### 2.1 Tablas verificadas (contratables desde el día uno)

**Bloque ASE1 golden (caché template) = resumen fuente ASE1 (verificado EXACTO):**

| Empresa | 1-APLICADOS A FACTURACION | 2-SALDOS A FAVOR | 3-FINANCIACIONES | 7-SERV. ESPECIALES | Total bloque |
|---|---|---|---|---|---|
| ENEL | 16194517449 | 50286069 | 0 | 0 | 16244803518 |
| OCCIDENTE | 523782724.81 | 12466967.19 | 4226000 | 0 | 540475692 |
| **Total ASE1** | — | — | — | — | **16785279210** |

**Mapa de columnas fuente por ASE (verificado, V10):**

| ASE | Columnas empresa en el Resumen | Occidente |
|---|---|---|
| ASE1 Promoambiental | ENEL | OCCIDENTE |
| ASE2 LIME | ENEL, NUEVO ESQUEMA | OCCIDENTE |
| ASE3 Ciudad Limpia | ENEL | OCCIDENTE |
| ASE4 Bogotá Limpia | ENEL, NUEVO ESQUEMA | — (verificar en T0: ¿OCCIDENTE ausente o cero?) |
| ASE5 Área Limpia | ENEL, ENERBIT | OCCIDENTE |

> Nota: la tabla V10 del orquestador asigna a ASE4 (Bogotá Limpia) el trío ENEL/NUEVO ESQUEMA/OCCIDENTE y a ASE3 (Ciudad Limpia) el dúo ENEL/OCCIDENTE, coherente con Apéndice B del Rector (ASE4 = ENEL + EAAB+CiudLimp; el "NUEVO ESQUEMA" es EAAB Reciprocidad del esquema actual — T0 confirma la etiqueta exacta en la fuente ASE4). Ninguna empresa ausente se inventa: ausente = 0 explícito solo si T0 demuestra columna inexistente; si la columna existe con 0, se lee el 0.

**Títulos consolidados filas 1–7 (V11):** ENEL (= facturación conjunta ENEL) · CIUDAD LIMPIA + EAB (= antiguo "EAAB + Ciud Limp" particionado por parte) · NUEVO ESQUEMA (= EAAB Reciprocidad, esquema actual) · ENERBIT · OCCIDENTE (= facturación directa Banco Occidente).

**C59 = quincena (`1` en Q1) + texto "OJO/// DILIGENCIAR". Filas 59–80 = "Recaudo {empresa} - ASE n" con TRUE/FALSE + diferencia (ej. ASE2 ENEL FALSE `397250` = an constructivo esperado, V9). Fila 80 "TOTAL RECAUDO" `58447133168`.**

(Tablas golden de bloques ASE2..ASE5 + clasificación valor-vs-fórmula de 1–7/59–80/TOTAL las extrae T0-0.5 del caché golden; este plan no las inventa.)

### 2.2 Decisiones de arquitectura

| ID | Opción elegida | Descartada | Por qué |
|---|---|---|---|
| D1 | Leer el Resumen por **búsqueda de etiqueta desde el final** del `Sheet1` ("Resumen Recaudo Aplicado Por Servicio", match normalizado por prefijo; filas 1/2/3/7 por prefijo `1-APLICADOS` / `2-SALDOS` / `3-APLICADOS A FINANCIACIONES` / `7-APLICADOS`) | Fila fija (r414–419 de ASE1) o string exacto | V4: las filas diarias varían por ASE; V6: variante sin "NUEVAS". Fila fija = bug garantizado en ASE2..5 |
| D2 | Filas 1–7 con **doble desenlace T0**: (a) si valores → `ReporteBancoInputs.Consolidado` se escribe con Σ de dominio; (b) si fórmulas → entran al mapa protegido y solo se verifica Σ | Presuponer valores o fórmulas antes de T0 | V1 (TRUE/TRUE en col J fila 7) impide presuponer. Ambos desenlaces usan el mismo gate Σ; solo cambia el mapa escribible |
| D3 | Misma pasada de escritura HU-07/HU-08 (overload lista; un `File.Copy` + validación pre/post con mapa ampliado) | Segundo pase / writer separado | La superficie nueva son más celdas leaf en la misma sesión. Un segundo pase rompería la atomicidad certificada y el hash A4 |
| D4 | `WorkbookLeafCellMapReporteBanco`: dict explícito (ASE × empresa → refs editables + visibles esperados), hermano de los mapas HU-07/HU-08 | Offsets aritméticos entre bloques o columnas fijas por empresa | Layouts heterogéneos de columnas por ASE (V10); riesgo §11.3 Rector |
| D5 | Gate doble en `ValidadorBasico`: (i) **bloque ASE = resumen fuente** ±0.5 por empresa×concepto; (ii) **consolidado 1–7 = Σ bloques** ±0.5 (aritmética de dominio en ambos desenlaces D2) | Validar solo totales o exigir igualdad con `Recaudo *` | V5 (exactitud ASE1) + V9 (diferencia 59–80 esperada ≠ 0): la igualdad con `Recaudo *` sería un falso fallo |
| D6 | Filas 59–80 + TOTAL (+ 1–7 si T0 dice fórmulas) al mapa de fórmulas protegidas; la diferencia E59:E80 se **reporta** en log, no se falla | Implementar la lógica 59–80 o ignorarla en la validación | Es contenido 2.7 (comparan contra hojas HU-08); 2.3 las protege y audita. Blindan 2.4–2.7 contra escritura accidental |
| D7 | Fuente única: `ReportePagosxBanco_*` por carpeta (prefijo nuevo en `ArchivoFuenteLocator`); C59 = `Periodo.NumeroQuincena` (dominio, no fuente) | Leer C59 de la fuente o pedirla en UI | La quincena es dato del período que el usuario ya selecciona; la fuente no la declara |
| D8 | Sin cambios funcionales UI: resumen por banco en `txtLog` + Serilog (una línea por ASE×empresa con esperado post-Excel + diferencia 59–80 como "informativa") | Grid por banco / selectores / restyle | Delta mínimo; OPA sigue siendo Ejecutar |

### 2.3 Escritura 2.3 (misma sesión atómica)

```text
File.Copy plantilla → salida (una vez, igual que HU-07/HU-08)
  └─► ValidarFormulasProtegidas (mapa HU-07/HU-08 + mapa 2.3: filas 59–80, TOTAL,
      1–7 si T0-dice-fórmulas, resto 2.4–2.7)
        └─► EscribirCeldasLeaf HU-07/HU-08 (bloques + empresa, intactos)
        └─► EscribirCeldasBanco 2.3 (bloques 9–58 en valores + C59 + consolidado 1–7 si T0-dice-valores)
              └─► revalidar fórmulas protegidas → guardar
```

Ante cualquier fallo: borrar salida parcial (patrón existente). Plantilla origen jamás mutada (hash A4 se mantiene y se extiende al mapa 2.3).

### 2.4 Dominio (Core, sin deps)

```csharp
public sealed class ReporteBancoEmpresaInputs  // una empresa dentro de un bloque ASE
{
    public string Empresa { get; set; } = "";   // "ENEL", "NUEVO ESQUEMA", "ENERBIT", "OCCIDENTE", "CIUDAD LIMPIA+EAB"
    public decimal AplicadosFacturacion { get; set; }      // concepto 1
    public decimal SaldosFavorGenerados { get; set; }      // concepto 2
    public decimal FinanciacionesNuevas { get; set; }      // concepto 3 (± NUEVAS)
    public decimal RecibosServEspeciales { get; set; }     // concepto 7 (Res. 27, incl. RCD)
    // Total = suma (propiedad calculada, análoga a ConsolidadoAse.TotalAse)
}

public sealed class ReporteBancoAseInputs  // bloque 10-filas de un ASE
{
    public Ase Ase { get; set; } = new();
    public IReadOnlyList<ReporteBancoEmpresaInputs> Empresas { get; set; } = ...;
}

public sealed class ReporteBancoInputs  // hoja completa
{
    public IReadOnlyList<ReporteBancoAseInputs> Ases { get; set; } = ...;  // 5 bloques
    public int Quincena { get; set; }  // → C59 (1 o 2; Q2 validado pero bloqueado aguas arriba)
    public IReadOnlyDictionary<string, decimal>? Consolidado { get; set; } // solo desenlace D2(a); null = D2(b)
}
```

`WorkbookLeafInputs` suma `ReporteBanco: ReporteBancoInputs?` (null = comportamiento HU-08 intacto; compatibilidad hacia atrás por construcción). `IWorkbookLeafInputReader`: método de lectura por banco (contrato intacto, implementación extendida). `ProcesadorPeriodo`: pasos 2.3 integrados al flujo por ASE (leer banco → leaf → validar gates D5) con fail-fast que nombra ASE **y empresa-columna**; una escritura al final (G4).

### 2.5 `IValidador` 2.3 (sin reabrir HU-04..08)

1. Todo lo HU-07/HU-08 intacto (matcheo estricto por `Ase.Id`, `AjustesSfT=0` Q1, `GranTotal=Σ`, gates leaf-vs-consolidado y Σ-empresas HU-08).
2. Nuevo (i): por cada ASE, empresa y concepto: `bloque template == resumen fuente` ±0.5; Total bloque = Σ conceptos ±0.5.
3. Nuevo (ii): consolidado 1–7 `== Σ bloques por empresa` ±0.5 (aritmética de dominio; vale en ambos desenlaces D2).
4. Nuevo (iii): C59 `== Periodo.NumeroQuincena`; si no, `CalculoInvalidoException`.
5. La diferencia filas 59–80 se reporta como informativa (V9); **MUST NOT** fallar por ella. Prohibida la confusión TotOpt-vs-banco: el validador nunca compara bloques 2.3 contra agregados HU-02.
6. El validador NO abre `.xlsx` (igual que HU-06/HU-07/HU-08).

### 2.6 UI — Visual Design Intent (delta mínimo)

Densidad Balanced, mismos GroupBoxes, sin restyle/colores/iconos. El `txtLog` agrega, por cada ASE, líneas banco (empresa × 4 conceptos + total, "esperado post-Excel") + Σ consolidado + C59 + nota "diferencias 59–80 informativas (anulado/reversado)". Serilog: mismos eventos HU-07/HU-08 con propiedad `Hoja = "REPORTE RECAUDO x BANCO"`. Sin nuevos controles.

### 2.7 Golden Capa A extendida a 2.3 (honestidad HU-06/HU-07/HU-08)

| # | Qué | Contra qué | Tol |
|---|---|---|---|
| A1 | Celdas banco escritas en la **salida** (bloques 9–58 + C59 + 1–7 si D2(a)) | Mismas celdas **leaf** del golden | ±0.5 |
| A2 | Visibles de **dominio** por ASE×empresa (bloque = fuente V5; Σ consolidado = Σ bloques) | Caché golden de bloques + filas 1–7 | ±0.5 |
| A3 | Filas 59–80 + TOTAL (+ 1–7 si D2(b)) siguen siendo fórmula en la salida | Estructura | n/a |
| A4 | SHA256 plantilla origen igual antes/después | — | n/a |
| A5 | **Prohibido** comparar caché de fórmula de la salida vs golden; **prohibido** usar `TotOpt` HU-02 o `Recaudo *` HU-08 como oráculo de bloques banco | — | prohibido |
| A6 | `Conciliaciones/` y `R*_Remuneracion` NO son oráculo | — | n/a |

Capa B (manual Excel: abrir, recalcular, comparar bloques + 1–7 + 59–80 vs golden) fuera de CI, protocolo §5.3.

### 2.8 File changes previstos

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Core/Models/ReporteBancoInputs.cs` | Crear | `ReporteBancoEmpresaInputs` + `ReporteBancoAseInputs` + `ReporteBancoInputs` (§2.4) |
| `Remuneracion.Core/Models/WorkbookLeafInputs.cs` | Modificar | Sumar `ReporteBanco` (nullable, default null) |
| `Remuneracion.Core/Interfaces/IWorkbookLeafInputReader.cs` | Modificar | Método lectura por banco (overload) |
| `Remuneracion.Core/Interfaces/IValidador.cs` | Modificar | Overload/gates D5 (bloque=fuente, Σ consolidado, C59) |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Modificar | Gates D5; diferencia 59–80 informativa, nunca fallo |
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` | Modificar | Pasos 2.3 por ASE, fail-fast nombra ASE+empresa-columna |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapReporteBanco.cs` | Crear | Mapa explícito ASE×empresa (congelado T0) |
| `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` | Modificar | Lectura Resumen desde el final, match normalizado (D1) |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs` | Modificar | Coherencia bloque=fuente + Σ consolidado |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Escribir celdas 2.3 en la misma pasada + mapa protegido ampliado |
| `Remuneracion.Infrastructure/FileSystem/ArchivoFuenteLocator.cs` | Modificar | Prefijo `ReportePagosxBanco_` (naming V13 verificado) |
| `Remuneracion.WinForms/Form1.cs` | Modificar | Resumen por banco en log |
| `Remuneracion.IntegrationTests/GoldenReporteBancoTests.cs` | Crear | Capa A 2.3 (§2.7) |
| `Remuneracion.IntegrationTests/ReporteBancoTests.cs` | Crear | Dominio: gates D5, label variants, C59, mismatch nombra ASE+columna |
| `Remuneracion.IntegrationTests/ProcesadorPeriodoTests.cs` | Modificar | Casos 2.3 (integración período con banco) |

**No tocar (salvo bug blocker):** `IPlantillaWriter`/validation-only (HU-04); `IRecaudoReader` agregados HU-02; coherencia `F25`-Extemp HU-05; semántica single-ASE y multi-ASE HU-07; mapa por empresa HU-08; `DetRetriRounder` (2.6); `requirements/` legado.

---

## 3. SPEC / Acceptance Criteria (trazabilidad Propuesta §10 + base docs)

### Requirement 1 — Lectura del Resumen por banco (CA-1; Detalle de plantilla Inst. banco; Proceso de Recaudo § banco)

El sistema **MUST** leer el "Resumen Recaudo Aplicado Por Servicio" del final de cada `ReportePagosxBanco` (búsqueda dinámica desde el final, match de conceptos por prefijo normalizado — tolera "FINANCIACIONES" ± "NUEVAS"). **MUST NOT** usar fila fija; si falta la etiqueta o una empresa-columna esperada → fallo que nombra ASE + empresa-columna, nunca valor inventado.

- GIVEN carpeta `1-Promoambiental` Q1 → WHEN lectura banco → THEN ENEL `16194517449/50286069/0/0` y OCCIDENTE `523782724.81/12466967.19/4226000/0` ±0.5 vs fuente.
- GIVEN empresa sin recaudo (columna inexistente demostrada por T0) → THEN ceros explícitos (no ausencia).

### Requirement 2 — Coherencia banco (CA-2)

El sistema **MUST** cumplir (i) bloque ASE = resumen fuente ±0.5 por empresa×concepto y (ii) consolidado 1–7 = Σ bloques ±0.5. **MUST NOT** presentar agregados HU-02 (TotOpt) ni valores `Recaudo *` HU-08 como valores banco.

- GIVEN leafs 2.3 Q1 → THEN tabla §2.1 exacta ±0.5 + equivalentes ASE2..ASE5 de T0-0.5.

### Requirement 3 — Escritura 2.3 preservando fórmulas (CA-3/CA-4; Prompt Maestro Vo paso valores)

El sistema **MUST** escribir solo celdas del mapa banco (bloques 9–58 + C59 + 1–7 si D2(a)) en la misma pasada HU-07/HU-08; **MUST NOT** escribir ninguna celda con `<f>` (incluye 59–80, TOTAL, 1–7 si D2(b), resto 2.4–2.7); **MUST NOT** mutar la plantilla; ante fallo **MUST** borrar la salida parcial.

- GIVEN 5 leafs banco válidos → WHEN overload lista → THEN cambian solo celdas del mapa; A3 verde.
- GIVEN mismatch en ASE3–OCCIDENTE → THEN `CalculoInvalidoException` que nombra ASE+empresa-columna, sin archivo certificado.

### Requirement 4 — C59 quincena (CA-3)

El sistema **MUST** escribir C59 = `Periodo.NumeroQuincena` (1 en Q1). **MUST NOT** leer C59 de ninguna fuente.

### Requirement 5 — Diferencias 59–80 informativas (CA-5; Detalle de plantilla E59:E80)

El sistema **MUST** reportar las diferencias 59–80 en log como informativas (anulado/reversado misma quincena, V9). **MUST NOT** fallar validación por ellas (el caso ASE2 ENEL FALSE `397250` es comportamiento esperado, no defecto).

### Requirement 6 — Trazabilidad y UI honesta (CA-6/CA-7)

Serilog + `txtLog` por ASE **y por empresa** (qué fuente→qué celdas, bloque vs fuente, Σ consolidado, C59). Resumen etiqueta "esperado post-Excel". Sin restyle.

### Requirement 7 — Golden Capa A 2.3 (CA-3/CA-5)

Matriz §2.7 para 5 ASE × empresas con insumos Q1. **MUST NOT** comparar caché de salida vs golden (A5).

| CA §10 | HU-09 |
|---|---|
| CA-1 | Lee el Resumen por banco de los 5 ASE (fail-fast nombra ASE+empresa-columna) |
| CA-2 | Bloque = fuente y consolidado = Σ ASE (agregados HU-02/HU-08 ≠ valores banco) |
| CA-3 | Capa A 2.3; filas 59–80 correctas post-Excel (Capa B manual residual) |
| CA-4 | Reassert mapa ampliado + plantilla no mutada (hash) |
| CA-5 | Gates D5 + coherencia leaf banco |
| CA-6 | Serilog + log por banco |
| CA-7 | Mismo flujo 5-ASE + resumen por banco |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 600–1000 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 T0 + mapa → PR2 lectura banco + locator → PR3 writer + mapa protegido → PR4 validador + coherencia + período/UI → PR5 golden + tests |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes. Chained PRs recommended: Yes. Chain strategy: pending.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 0 | Discovery T0 + mapa congelado | PR 1 | Bloquea todo; solo lectura + datos |
| 1 | Lectura banco + locator | PR 2 | Depende de PR 1 |
| 2 | Writer 2.3 + protegido ampliado | PR 3 | Depende de PR 1 |
| 3 | Validador D5 + período + UI | PR 4 | Depende de PR 1 |
| 4 | Golden 2.3 + tests + Capa B doc | PR 5 | Depende de PR 2–4; 33 tests previos verdes |

### Phase 0 — Discovery T0 (BLOQUEANTE, solo lectura)

- [ ] 0.1 Clasificar valor-vs-fórmula celda por celda: filas 1–7 (con foco en col J fila 7 TRUE/TRUE), bloques 9–58, C59, filas 59–80, fila 80 TOTAL. Registrar fórmula visible de cada celda con `<f>`.
- [ ] 0.2 Localizar el "Resumen Recaudo Aplicado Por Servicio" en las fuentes ASE2..ASE5 (etiqueta exacta, distancia al final, headers de empresa-columna y orden, filas de conceptos 1/2/3/7: ¿las 4 siempre presentes?).
- [ ] 0.3 Confirmar mapa de columnas por ASE (V10 + etiqueta exacta de NUEVO ESQUEMA en ASE2/ASE4; OCCIDENTE en ASE4: ¿columna ausente o cero?; ENERBIT solo ASE5).
- [ ] 0.4 Registrar variantes de etiqueta (FINANCIACIONES ± NUEVAS y cualquier otra) y fijar la normalización (lowercase, sin "nuevas", por prefijo `1-aplicados` / `2-saldos` / `3-aplicados a financiaciones` / `7-aplicados`).
- [ ] 0.5 Extraer del caché golden las tablas de bloques ASE2..ASE5 (equivalente a §2.1) + verificar `GERENTES_*`/2.4–2.7 como fórmulas + barrido de externalLinks en la hoja banco.
- [ ] 0.6 Determinar semántica de C59 (¿la referencia alguna fórmula? ¿qué rango acepta?) y el veredicto D2(a)/D2(b) para filas 1–7.
- [ ] 0.7 Congelar `WorkbookLeafCellMapReporteBanco` (ASE × empresa → editables + visibles + C59 + consolidado según D2). **Nada entra al código sin esta tabla.**

### Phase 1 — Dominio (modelos + lectura)

- [ ] 1.1 `ReporteBancoInputs.cs` (§2.4; Total como propiedad calculada) + extender `WorkbookLeafInputs` (nullable, default null).
- [ ] 1.2 Reader banco según T0-0.7 (búsqueda desde el final, match normalizado, `ServEspK`-análogo n/a); fallo nombra ASE+empresa-columna.
- [ ] 1.3 `ArchivoFuenteLocator`: prefijo `ReportePagosxBanco_` (V13).

### Phase 2 — Escritura (misma pasada)

- [ ] 2.1 `WorkbookLeafCellMapReporteBanco.cs` (datos T0-0.7) + escritura de celdas 2.3 dentro del overload lista existente.
- [ ] 2.2 Mapa protegido ampliado (59–80, TOTAL, 1–7 si D2(b), resto 2.4–2.7; shared-formula awareness HU-07T0 §0.5).
- [ ] 2.3 Borrado de parcial + hash plantilla intactos.

### Phase 3 — Orquestación + UI delta mínimo

- [ ] 3.1 `ProcesadorPeriodo`: pasos 2.3 por ASE (leer banco → leaf → validar D5), fail-fast ASE+empresa-columna, una escritura; C59 desde `Periodo.NumeroQuincena`.
- [ ] 3.2 `Form1`: resumen por banco en log (§2.6).
- [ ] 3.3 Serilog por banco (propiedad `Hoja`, mismos sinks).

### Phase 4 — Pruebas y evidencia

- [ ] 4.1 `ReporteBancoTests` (in-memory + fuentes Q1 reales sin salida): Σ ok, label variant sin-NUEVAS matchea, C59 ≠ quincena falla, mismatch nombra ASE+columna, TotOpt-vs-banco prohibido (test que intenta la confusión y falla).
- [ ] 4.2 `GoldenReporteBancoTests`: matriz §2.7 (5 ASE × empresas, fixtures Q1).
- [ ] 4.3 Integración período con banco + negativa (falta etiqueta Resumen en fuente ASE3 → ASE+columna en el error, sin salida).
- [ ] 4.4 Los 33 tests existentes verdes; build 0 warnings; CRLF; sin commit.

### Phase 5 — Documental

- [ ] 5.1 Capa B manual §5.3 ejecutada una vez y evidenciada (sin fingirla como gate de merge).
- [ ] 5.2 Cierre deja explícito el frente 2.4 (siguiente HU propuesta: `BCE SC POR FACT.`).

---

## 5. Test Strategy

| Capa | Qué | Cómo |
|---|---|---|
| Unidad | Gates D5 (bloque=fuente, Σ consolidado, C59) | In-memory, tabla §2.1 + T0-0.5 |
| Unidad | Reader banco (búsqueda final, label variants, mapa por ASE LIME 3-col vs PROMO 2-col) | Fuentes Q1 reales, sin Excel de salida |
| Integración | Período con 2.3 (5 carpetas reales, salida temp) | Insumos Q1, fail-fast ASE+empresa-columna |
| Golden Capa A 2.3 | Matriz §2.7 | OpenXML read-only + aritmética dominio; plantilla=golden (A5 aplica) |
| UI | Resumen por banco | Funcional manual (sin harness) |
| Capa B | Bloques + 1–7 + 59–80 post-Excel | Manual — §5.3 |

### 5.1 Fixtures

- Golden/plantilla: `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` (hoja `REPORTE RECAUDO x BANCO`).
- Fuentes: `Docs/Insumos/REMUNERACION 2026071/{1..5}-*/ReportePagosxBanco_*.xlsx` (Q1; Q2 `REMUNERACION 2026072/` bloqueada — sigue lanzando `CalculoInvalidoException` aguas arriba).
- Referencia: tabla §2.1 (ASE1) + T0-0.5 (ASE2..ASE5), tolerancia ±0.5.

### 5.2 Casos negativos obligatorios (nombran ASE y empresa-columna)

Resumen sin etiqueta final en ASE3; variante "FINANCIACIONES" sin NUEVAS debe matchear (no fallar); columna ENERBIT pedida en ASE1 (inexistente → 0 solo si T0 lo autoriza, si no fallo); Σ OCCIDENTE ≠ fuente en ASE5; C59 ≠ quincena; salida == plantilla (no in-place); overwrite cancelado; Q2 sigue bloqueada; TotOpt HU-02 usado como valor banco (prohibido — el test lo demuestra fallando).

### 5.3 Protocolo manual Capa B (no CI)

1. Generar salida a ruta distinta del golden. 2. Abrir en Excel, recalcular. 3. Comparar bloques 9–58 + 1–7 + C59 + 59–80 vs golden ±0.5 (refs §2.1 + T0-0.5) y verificar que 59–80/TOTAL calcularon (ej. ASE2 ENEL FALSE `397250`). 4. Evidencia adjunta. No es gate de merge.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | Banco = extensión de datos (mapa + inputs + reader + gates); período/orquestación/writer conservan su rol HU-07/HU-08. |
| **O** | Se agregan modelos/mapas/overloads; el path HU-08 funciona con `ReporteBanco = null` (abierto sin modificar). |
| **L** | `OpenXmlPlantillaWriter` escribe más celdas del mismo modo leaf; comportamiento HU-08 puro inalterado. |
| **I** | `ReporteBancoInputs` separado de `ConciliacionEmpresaInputs`; reader/validador crecen por overload. |
| **D** | Core define banco/inputs/gates; Infrastructure/WinForms componen. Sin nuevas deps. |

### 6.2 Best Practices

- La verdad del workbook manda: filas 59–80/TOTAL (y 1–7 si T0 dice fórmulas) ⇒ cero escritura directa (Rector §6 preservación).
- Mapa explícito ASE×empresa verificado, no offsets ni filas fijas (Rector §11.3; V4/V10).
- Una escritura atómica; plantilla nunca mutada; hash A4 extendido.
- Golden honesto por banco (OpenXML no recalcula; A5/A6).
- Fail-fast nombra ASE+empresa-columna; sin salida certificada ante fallo.
- Documentos base citados por sección como spec funcional (§1.5), no como folklore.
- Tándem HU-08: 59–80 se protege aquí y se implementa en 2.7; HU-09 no duplica ni anticipa 2.7.

### 6.3 Performance

- 5 bloques × ~10 celdas + C59 + consolidado (si D2(a)) ≈ <60 escrituras extra en la misma sesión OpenXML + 5 lecturas de Resumen (búsqueda desde el final sobre ~400 filas, ExcelDataReader streaming). Irrelevante a esta escala; `Task.Run` existente para no congelar el form.

**Veredicto:** APROBADO como it. 2.3 de Fase 2 **si** T0 congela el mapa con evidencia (incluido el veredicto D2 filas 1–7) y se acepta CA-3 parcial (Capa A en CI, Capa B manual).

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| Filas 1–7 son fórmulas no editables y el plan asumía valores (o viceversa) | Media | D2 doble desenlace: el gate Σ es idéntico; solo cambia el mapa escribible. T0-0.1/0.6 lo decide antes de codificar |
| El Resumen de ASE2..ASE5 trae etiquetas/columnas distintas a ASE1 | Alta | D1 (búsqueda + normalización) + T0-0.2/0.3; si el hueco es real, la HU se recorta a los ASE verificados con rebase, no se inventa |
| Algún concepto 1/2/3/7 falta como fila cuando vale 0 | Media | T0-0.2; el reader distingue "fila ausente = 0 demostrable" de "columna ausente = fallo" (regla fijada en T0-0.7) |
| C59 es referenciada por fórmulas con semántica distinta a "quincena" | Baja | T0-0.6; si C59 alimenta lógica 2.7, se escribe igual (es valor) y se documenta el efecto |
| Inflar a 2.4–2.7 dentro de esta HU (BCE tienta, 59–80 tienta) | Media | §0.2 out-of-scope + D6 (protegidas, no implementadas); rechazar PRs que lo metan |
| Comparar caché de salida vs golden y "cerrar" CA-3 en falso | Alta | A5 en tests; este plan lo prohíbe |
| Carpeta `5-Área Limpia` (tilde) en nuevos tests | Media | Reutilizar `Insumos.cs` (ya lo maneja) |
| Q2 (`REMUNERACION 2026072/`) confundida con fixture Q1 | Media | Fixtures §5.1 fijan Q1; Q2 sigue bloqueada aguas arriba |

---

## 8. Rollback

- Revertir PRs en orden inverso (golden → período/UI → validador → writer → reader/locator → mapa/T0).
- HU-04..HU-08 intactas sin esta HU: `ReporteBanco = null` = comportamiento HU-08 puro.
- `Docs/Insumos/` untracked: jamás como destino; pisada se restaura desde backup.
- Si T0 demuestra hueco de fuente (Riesgo 2), recorte con rebase, no invención.

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §9 it. 2.3 / §6 / §10 (±0.5) + tres documentos base como origen funcional (§1.5). 2.4–2.7 y Q2 quedan fuera, citados a §9.
2. Lectura del Resumen por búsqueda dinámica desde el final con match normalizado (V4/V6); fila fija y string exacto prohibidos.
3. Filas 59–80 + TOTAL = fórmulas protegidas (hipótesis V8, T0 confirma); su diferencia es informativa, nunca fallo (V9). Su lógica es 2.7, no 2.3.
4. Filas 1–7 con doble desenlace T0 (D2): valores → se escriben con Σ de dominio; fórmulas → se protegen y solo se verifica Σ.
5. Una sola escritura atómica (overload lista HU-07/HU-08); sin segundo pase.
6. Gates D5 (bloque=fuente, Σ consolidado, C59=quincena) ±0.5; fail-fast nombra ASE+empresa-columna.
7. Mapa explícito ASE×empresa congelado por T0; ninguna dirección sin evidencia.
8. Golden Capa A 2.3 con honestidad HU-06/HU-07/HU-08; Capa B manual residual.
9. UI delta mínimo + Serilog por banco; OPA = Ejecutar.
10. Apply espera aprobación + PRs encadenados (Unidad 0–4, §4).

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
