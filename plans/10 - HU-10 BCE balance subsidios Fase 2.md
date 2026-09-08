# Plan 10 — HU-10: BCE SC POR FACT. Balance Subsidios y Contribuciones (Fase 2, it. 2.4)

> **Historia:** diligenciar la hoja `BCE SC POR FACT.` para los 5 ASE desde la fila "Total General" de cada `R4-BalanceSubsidioyContribuciones_*.xlsx`, dentro de la misma escritura atómica multi-ASE certificada por HU-07/HU-08/HU-09.
> **Rector (IRRENUNCIABLE):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 2 it. **2.4 BCE SC POR FACT**, §6 arquitectura (capas, preservación de fórmulas, Serilog, lectura directa de fuente), §10 CA + tolerancia ±0.5. **Entra SOLO 2.4.** Salen 2.5 AJUSTES-SF-T, 2.6 DetRetri, 2.7 validaciones cruzadas — son HUs posteriores. Quincena 2 sigue bloqueada.
> **Origen funcional:** los tres documentos base — `Detalle de plantilla.docx` (Inst. hoja `BCE SC POR FACT.`), `Proceso de Recaudo.docx` (§ balance de subsidios y contribuciones por facturación), `Prompt Maestro Vo.docx` (paso de pegado en valores del Balance) — citados por sección en §1.5/§3.
> **Continuidad:** HU-01..HU-09 cerradas o en cierre (ruta leaf multi-ASE atómica single-write, mapas explícitos por `Ase.Id`, validador estricto, UI modo 5-ASE, Golden Capa A, build 0/0, tests 47/47, harness 24/24). Este plan NO reabre su semántica. Tándem con HU-09: mismo patrón de extensión leaf (mapa + inputs + reader + gates + golden).
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-08

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. El orquestador ya verificó el golden Q1 (`Remuneracion 202607-1 Total.xlsx`, hoja `BCE SC POR FACT.` = sheet9, 40 sheets mapeadas sheet1..sheet40) y la fuente Q1 ASE1 contra el template; lo que queda abierto es (a) la **contradicción D/E** entre el texto del documento base y los headers del template, (b) el carácter valor-vs-fórmula de C2:E7, F2:F7, H2:H7, filas 9/11 y bloque 18–24, y (c) la ubicación exacta del "Total General" en las fuentes ASE2..ASE5 (incluida la variante Optimizado de ASE5) — por eso el plan incluye discovery T0 obligatorio y bloqueante (§4 Fase 0). El apply espera aprobación explícita.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (dump OpenXML del golden Q1 sheet9 + lectura fuente ASE1 Q1 real, reportado por el orquestador):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | Template `BCE SC POR FACT.` (sheet9): header fila 2 = C:"ASE", D:"CONTRIBUCION", E:"SUBSIDIO", F:"TOTAL BSC", H:"SISTEMA" | Dump raw golden (orquestador) | Los headers del template son la hipótesis líder (D=contribución positiva, E=subsidio negativo); T0 debe probarla contra la fuente |
| V2 | Filas 3–7 por ASE: ASE id + D(+) + E(−) + F(=D+E) + H(redondeado?) + I(diff ~±0.4). ASE1: D=`3256235169.97`, E=`-1223871491.13`, F=`2032363678.84`, H=`2032363679`, I=`-0.16` | Caché golden | Oráculo Capa A para ASE1; F=D+E verificable en dominio; H≈F con regla de redondeo que fija T0 |
| V3 | Fila 9: F total `-5350444198.33` + H/I. Fila 11: sumas de columnas D/E/F | Caché golden | Filas 9/11 candidatas a **fórmulas protegidas** (T0 clasifica valor-vs-fórmula celda por celda) |
| V4 | Filas 18–24 = bloque de validación (ASE, REMUN TOTAL, VALIDA 1=TRUE, VALIDA 2=FALSE con diffs ±0.16–0.41 — dentro de tolerancia ±0.5) | Caché golden | Bloque 18–24 candidato a **fórmulas protegidas** (territorio 2.7); T0 lo confirma |
| V5 | **Blast radius verificado (fórmulas volcadas):** CONSOLIDADO (sheet18) J9:J13 = `BCE SC POR FACT.!F3:F7` por ASE; K9 = D9-E9-F9-G9-H9-I9-J9; M9 = K9-L9 (L9 = INTERVENTORIA!Q4). DetRetri (sheet36) + DetValiRetri (sheet37) referencian `-BCE!F3..F7` dentro de fórmulas grandes de validación | Dump de fórmulas (orquestador) | 2.4 escribe BCE C2:E7 (+H si T0 lo autoriza) y J9:J13 + K/M downstream **calculan solos, nunca se escriben**. DetRetri/DetValiRetri = 2.6/2.7 → protegidas, nunca escritas |
| V6 | Fuente ASE1: secciones por localidad (USAQUEN/CHAPINERO/…) con filas "Total"; fila **"Total General" al FINAL** (posición varía → búsqueda dinámica, nunca fila fija) | Lectura fuente real | Estrategia de lectura: **búsqueda dinámica de etiqueta "Total General" desde el final** |
| V7 | Columnas del reporte fuente: B=productor, C=usuarios-res, D=usuarios-no-res, E=Subsidio, F=Contribución, G=Valor | Lectura fuente real | Base de la resolución D/E: la prueba es E-fuente vs F-fuente contra D/E-template (ver §0.2) |
| V8 | Naming confirmado en disco por carpeta ASE: `R4-BalanceSubsidioyContribuciones_to_date...xlsx`; variante ASE5: `R4-BalanceSubsidioyContribuciones-Optimizado_...xlsx` | `Get-ChildItem` Insumos | El locator debe aceptar **ambos prefijos** (`R4-BalanceSubsidioyContribuciones` y `...-Optimizado_`); sin nuevos tipos de archivo |
| V9 | Rango funcional según doc base: C2:E7 de la hoja (columna D ← subsidio / columna E ← contribución **según TEXTO DEL DOC — ver contradicción**); F2:F7 = subsidio+contribución; fila 16+ = validaciones del proceso; el Balance final alimenta CONSOLIDADO_TOTAL RECAUDO | Documentos base (orquestador) | El rango C2:E7 + F2:F7 + validaciones es el alcance funcional; la asignación D/E del texto se somete a T0 (no se codifica sin prueba) |
| V10 | Contratos listos para extender: overload lista multi-ASE (`IWorkbookLeafWriter`), `WorkbookLeafInputs` con listas, `Insumos.cs` + `CarpetasAse.Prefijos` cubren las 5 carpetas; `ArchivoFuenteLocator` resuelve por prefijo | Lectura de código | 2.4 = **extensión del patrón HU-07/HU-08/HU-09**, no arquitectura nueva |
| V11 | Quincena 2 sigue bloqueada aguas arriba (`CalculoInvalidoException`); `AjustesSfT=0` en Q1 intacto | HU-05/HU-07 | 2.4 no toca Q2 ni AJUSTES-SF-T |

**NO verificado (y por eso T0 es bloqueante, §4 Fase 0):**

1. **La contradicción D/E:** el texto del doc base dice D=subsidio, E=contribución; los headers del template dicen D=CONTRIBUCION, E=SUBSIDIO y los valores golden concuerdan con los HEADERS (D positivo, E negativo). T0 debe leer la fila "Total General" de la fuente (E-fuente=Subsidio? F-fuente=Contribución?) y **probar** qué columna fuente alimenta template-D vs template-E. Nada de D/E entra al código sin esta prueba.
2. Carácter valor-vs-fórmula celda por celda de BCE C2:E7, F2:F7, H2:H7 (¿es H valor pegado o fórmula `=ROUND(F,0)`?), filas 9/11, bloque 18–24.
3. Ubicación y etiqueta exacta del "Total General" en las fuentes ASE2..ASE5 (¿misma etiqueta? ¿misma posición relativa al final? ¿variantes con espacios/mayúsculas?).
4. Paridad de layout de la variante Optimizado de ASE5 (¿mismas columnas B..G? ¿misma etiqueta "Total General"? ¿misma hoja?).
5. Semántica exacta de columna H (SISTEMA): ¿valor pegado desde dónde, o fórmula de redondeo? ¿columna I = diferencia calculada (fórmula) o valor?
6. Si alguna celda objetivo de 2.4 depende de un externalLink roto.
7. Tablas golden de filas 3–7 para ASE2..ASE5 (las extrae T0-0.5 del caché; este plan no las inventa).

> **Regla de hierro del plan:** ninguna dirección de celda de BCE entra al código sin pasar por T0, y la asignación D/E no se codifica por hipótesis sino por prueba fuente-vs-template. Lo ya verificado arriba (V1–V11) sí es contratable desde el día uno.

### 0.2 La contradicción D/E — estrategia de resolución (T0)

| Fuente de verdad | Qué dice | Peso |
|---|---|---|
| Texto del doc base (V9) | Rango C2:E7: columna D ← subsidio, columna E ← contribución | Bajo sin prueba — el texto puede tener D/E cruzados |
| Headers del template (V1) | D=CONTRIBUCION, E=SUBSIDIO | Alto — es lo que el usuario ve y valida |
| Valores golden (V2) | D positivo (`3256235169.97`), E negativo (`-1223871491.13`), F=D+E | Alto — la aritmética D+E=F cierra con los headers |
| Columnas fuente (V7) | E-fuente=Subsidio, F-fuente=Contribución | Alto — es el origen |

**Hipótesis líder (a probar, no a asumir):** template-D (CONTRIBUCION) ← F-fuente (Contribución); template-E (SUBSIDIO) ← E-fuente (Subsidio). Es decir, el texto del doc base tendría D/E invertidos respecto al template.

**Protocolo T0 (§4 Fase 0, tarea 0.2):** para ASE1 (y luego ASE2..ASE5): (i) leer la fila "Total General" de la fuente con ExcelDataReader (valores, no caché); (ii) extraer E-fuente y F-fuente; (iii) comparar contra template D3/E3 del golden: el match ±0.5 decide la asignación; (iv) si E-fuente ≈ E-template y F-fuente ≈ D-template → hipótesis líder **probada** y se congela así en el mapa; (v) si el match es el inverso → se congela la asignación del texto del doc y se documenta que los headers/golden se leen cruzados; (vi) si **ningún** match cierra ±0.5 → T0 se declara **fallido**, la HU se recorta y se escala al Ingeniero (no se inventa mapeo). El veredicto queda registrado en `07 - HU-07 T0 Evidencia.md` (anexo) o nota T0 propia según convención vigente.

### 0.3 Mapeo al Rector (in vs out)

**Entra porque §9 it. 2.4 / §6 / §10 lo piden ahora:**

| Rector | Qué cubre HU-10 |
|---|---|
| §9 it. 2.4 | `BCE SC POR FACT.`: C2:E7 (+H según T0) + F2:F7 como protegidas + filas 9/11 y 18–24 como protegidas |
| §5 fila Balance SC | Subsidio / contribución desde `R4-BalanceSubsidioyContribuciones` (col. E/F fuente según T0) |
| §7.1 bloque escritura | "Balance SC → pega subsidio/contribución" (valores, misma pasada atómica) |
| §10 CA-1/CA-2 | Leer el Total General de los 5 ASE + coherencia BCE = fuente y F=D+E |
| §10 CA-4/CA-5/CA-6/CA-7 | Fórmulas intactas (mapa ampliado 2.4), validación por ASE, Serilog por ASE, UI sin cambios salvo resumen |

**Sale porque §9 lo asigna a 2.5–2.7 / Fase 3 (HUs posteriores):**

- 2.5 AJUSTES-SF-T (`SALDOS POR NOTA` + `RETRIBUCION NEGATIVA`; Q1 sigue con `AjustesSfT=0`; quincena 2 sigue bloqueada en `CalculoRemuneracion`).
- 2.6 `DetRetri2026071` / `DetValiRetri2026071` como objetivo de escritura (sus fórmulas que referencian `-BCE!F3..F7` se protegen; su lógica no se implementa).
- 2.7 validaciones cruzadas (`VALIDACION_TOTAL`, `VALIDACION_RECIP/ENEL/…`, bloque BCE 18–24 más allá de protegerlo, CONSOLIDADO J/K/M más allá de verificar cálculo automático). Se protegen sus fórmulas; su lógica no se implementa.
- `INTERVENTORIA` (L9 alimenta M9 — fuera de alcance), `REPORTE RECAUDO x BANCO` (HU-09, intacto), resto de hojas Fase 2/3.

### 0.4 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **T0 bloqueante con protocolo de resolución D/E** (§0.2). Sin veredicto probado fuente-vs-template no se escribe ni una celda BCE. |
| G2 | **Lectura dinámica del "Total General" desde el final** (búsqueda de etiqueta normalizada). Prohibida fila fija: la posición varía por ASE (V6). |
| G3 | **F2:F7 + filas 9/11 + bloque 18–24 como fórmulas protegidas** (hipótesis V3/V4; T0 la confirma o corrige). Esta HU no implementa su lógica (es 2.7). |
| G4 | **Columna H: T0 decide valor-vs-fórmula.** Si es valor → se escribe desde dominio (redondeo definido por T0). Si es fórmula (`=ROUND(F,0)` u otra) → se protege y solo se verifica H≈F. El plan soporta ambos desenlaces sin rebase (ver D2). |
| G5 | **Una sola llamada de escritura** por proceso (mismo overload lista HU-07/HU-08/HU-09). Sin segundo pase (ver D3). |
| G6 | **Modelo explícito por ASE** (`BalanceScInputs`: Subsidio, Contribucion, TotalBSC por `Ase.Id`), no offsets. Columnas fuente heterogéneas solo en la variante Optimizado (V8) — T0 fija paridad. |
| G7 | **Downstream automático:** CONSOLIDADO J9:J13 + K/M calculan por fórmulas (V5); el validador los verifica post-Excel en Capa B, nunca los escribe ni los exige en Capa A. |
| G8 | **UI sin cambios funcionales**: el modo 5-ASE ya existe; 2.4 solo agrega líneas de resumen BCE al log. Sin restyle. |
| G9 | **Golden Capa A extendida a 2.4** con honestidad HU-06..HU-09 (leaf salida vs leaf golden; visibles de dominio vs caché golden; nunca caché de salida vs golden). |

---

## 1. PROPOSE

### 1.1 Intent

Diligenciar la hoja `BCE SC POR FACT.` para los 5 ASE dentro de la misma escritura atómica del período — C2:E7 en valores desde el "Total General" de cada Balance (asignación D/E probada por T0, no asumida), H2:H7 según veredicto T0, F2:F7 + filas 9/11 + bloque 18–24 como fórmulas protegidas — de modo que CONSOLIDADO J9:J13 + K/M downstream calculen solos y la coherencia BCE = fuente quede demostrada contra el golden Q1.

### 1.2 In Scope

- Discovery T0 (§4 Fase 0) + `WorkbookLeafCellMapBalanceSc` congelado solo con evidencia (incluido el veredicto D/E y el veredicto H).
- Lectura del "Total General" desde las 5 carpetas (prefijos `R4-BalanceSubsidioyContribuciones` y `...-Optimizado_`).
- Extensión de `WorkbookLeafInputs` + escritura en la misma pasada HU-07/HU-08/HU-09; validación pre/post con mapa ampliado (F2:F7, filas 9/11, bloque 18–24, CONSOLIDADO J/K/M y DetRetri/DetValiRetri como protegidas).
- Gate BCE = fuente por ASE y gate F=D+E en `IValidador` (+ coherencia leaf).
- Golden Capa A extendida a 2.4 + tests con insumos Q1 reales.
- Resumen BCE por ASE en log/Serilog (delta mínimo UI).

### 1.3 Out of Scope

Todo §0.3 (2.5–2.7, Fase 3). Además, explícito: **escrituras en DetRetri/DetValiRetri** (aunque referencien `-BCE!F3..F7`, son 2.6/2.7); **lógica de INTERVENTORIA** (L9 alimenta M9 — no se toca); quincena 2 (sigue lanzando `CalculoInvalidoException`); reescritura del path HU-07/HU-08/HU-09 (se extiende, no se duplica); bloque 18–24 como lógica implementada (solo protegido); restyle UI; DI framework.

### 1.4 Resultado de negocio

El Ingeniero ejecuta el modo 5 ASE como hoy; obtiene el mismo workbook más la hoja `BCE SC POR FACT.` diligenciada (C2:E7 en valores probados contra fuente, H según T0, F + totales + validaciones calculando por fórmulas); CONSOLIDADO J9:J13 refleja los TOTAL BSC automáticamente; el log audita por ASE **subsidio/contribución/total**; las pruebas demuestran BCE = fuente y F=D+E.

### 1.5 Base documental (origen funcional — citas por sección)

| Documento | Sección / instrucción | Qué aporta a 2.4 |
|---|---|---|
| `Detalle de plantilla.docx` | Inst. hoja `BCE SC POR FACT.` (rango C2:E7 por ASE, F2:F7 = subsidio+contribución, fila 16+ validaciones del proceso) | Layout de la hoja y semántica de cada rango; base del mapa T0. **Nota:** su asignación textual D=subsidio/E=contribución se somete a T0 por contradicción con headers (§0.2) |
| `Proceso de Recaudo.docx` | § balance de subsidios y contribuciones por facturación (origen: reporte Balance por ASE, periodicidad quincenal) | Origen de los valores: el "Total General" del reporte Balance por ASE resume subsidio y contribución de la quincena |
| `Prompt Maestro Vo.docx` | Paso de pegado en valores del Balance SC por ASE + destino CONSOLIDADO_TOTAL RECAUDO | La hoja se pega en valores (celdas editables C2:E7); el Balance final alimenta el consolidado (vía J9:J13 por fórmulas, V5) |
| Rector Propuesta | §9 it. 2.4 (alcance), §6 (preservación de fórmulas, lectura directa, trazabilidad), §10 CA + tolerancia ±0.5, §5 fila Balance SC | Rector normativo |

---

## 2. DESIGN

### 2.1 Tablas verificadas (contratables desde el día uno)

**Fila ASE1 golden (caché template) — oráculo Capa A para ASE1:**

| ASE | D (CONTRIBUCION, +) | E (SUBSIDIO, −) | F (=D+E) | H (SISTEMA) | I (diff) |
|---|---|---|---|---|---|
| ASE1 | `3256235169.97` | `-1223871491.13` | `2032363678.84` | `2032363679` | `-0.16` |

**Totales y validación golden:**

| Rango | Valor / estado |
|---|---|
| Fila 9, col F (total) | `-5350444198.33` + H/I |
| Fila 11 (sumas D/E/F) | sumas de columnas (protegidas según T0) |
| Filas 18–24 (validación) | VALIDA 1=TRUE, VALIDA 2=FALSE, diffs ±0.16–0.41 (dentro de ±0.5) |

**Blast radius (fórmulas volcadas, V5 — verificado):**

| Hoja | Celda | Fórmula | Tratamiento 2.4 |
|---|---|---|---|
| CONSOLIDADO (sheet18) | J9:J13 | `=BCE SC POR FACT.!F3:F7` por ASE | Protegidas; verificadas en Capa B post-Excel, nunca escritas |
| CONSOLIDADO | K9 | `=D9-E9-F9-G9-H9-I9-J9` (y análogas) | Protegidas; cálculo automático |
| CONSOLIDADO | M9 | `=K9-L9` (L9 = INTERVENTORIA!Q4) | Protegida; INTERVENTORIA fuera de alcance |
| DetRetri (sheet36) | refs | `-BCE!F3..F7` en fórmulas grandes | Protegidas (2.6); nunca escritas |
| DetValiRetri (sheet37) | refs | `-BCE!F3..F7` en fórmulas grandes | Protegidas (2.7); nunca escritas |

**Columnas fuente (verificado, V7):** B=productor, C=usuarios-res, D=usuarios-no-res, E=Subsidio, F=Contribución, G=Valor. La asignación a template-D/E la prueba T0 (§0.2).

**Naming fuente (verificado, V8):** prefijo base `R4-BalanceSubsidioyContribuciones` + variante ASE5 `R4-BalanceSubsidioyContribuciones-Optimizado_`. El locator acepta ambos.

> Tablas golden de filas 3–7 para ASE2..ASE5 + clasificación valor-vs-fórmula de C2:E7/F2:F7/H2:H7/filas 9/11/bloque 18–24 las extrae T0-0.5 del caché golden; este plan no las inventa.

### 2.2 Decisiones de arquitectura

| ID | Opción elegida | Descartada | Por qué |
|---|---|---|---|
| D1 | Leer el "Total General" por **búsqueda de etiqueta desde el final** del `Sheet1` (match normalizado: lowercase, trim, sin espacios múltiples; etiqueta canónica `"total general"`) | Fila fija (posición de ASE1) o string exacto con mayúsculas | V6: las secciones por localidad varían por ASE; fila fija = bug garantizado en ASE2..5 |
| D2 | Columna H con **doble desenlace T0**: (a) si valor → `BalanceScInputs` incluye `Sistema` y se escribe con redondeo fijado por T0; (b) si fórmula → entra al mapa protegido y solo se verifica H≈F | Presuponer valor o fórmula antes de T0 | V2 (H=`2032363679` ≈ ROUND(F,0)) impide presuponer. Ambos desenlaces usan el mismo gate H≈F; solo cambia el mapa escribible |
| D3 | Misma pasada de escritura HU-07/HU-08/HU-09 (overload lista; un `File.Copy` + validación pre/post con mapa ampliado) | Segundo pase / writer separado | La superficie nueva son ~15 celdas BCE en la misma sesión. Un segundo pase rompería la atomicidad certificada y el hash A4 |
| D4 | `WorkbookLeafCellMapBalanceSc`: dict explícito (ASE → refs editables C/D/E(+H) + visibles esperados), hermano de los mapas HU-07/HU-08/HU-09 | Offsets aritméticos o columnas fijas sin mapa | Riesgo §11.3 Rector; la asignación D/E depende del veredicto T0 y debe quedar congelada en un solo lugar auditable |
| D5 | Gate doble en `ValidadorBasico`: (i) **BCE por ASE = Total General fuente** ±0.5 (con la asignación D/E del veredicto T0); (ii) **F=D+E** ±0.5 por ASE + sumas de fila 11 en dominio | Validar solo totales o exigir igualdad con CONSOLIDADO J/K/M en Capa A | V5: J/K/M son fórmulas que OpenXML no recalcula — exigirlas en Capa A sería un falso fallo (van a Capa B) |
| D6 | F2:F7 + filas 9/11 + bloque 18–24 + J/K/M + refs DetRetri/DetValiRetri al mapa de fórmulas protegidas | Implementar la lógica 18–24 o ignorarla en la validación | Es contenido 2.6/2.7; 2.4 las protege y audita. Blindan 2.5–2.7 contra escritura accidental |
| D7 | Fuente única: `R4-BalanceSubsidioyContribuciones_*` por carpeta (dos prefijos en `ArchivoFuenteLocator`); sin dato de UI | Pedir subsidio/contribución en UI o leerlos de otra hoja | Son datos de la fuente quincenal por ASE; la UI ya selecciona período y carpetas |
| D8 | Sin cambios funcionales UI: resumen BCE en `txtLog` + Serilog (una línea por ASE con subsidio/contribución/total + veredicto D/E citado) | Grid BCE / selectores / restyle | Delta mínimo; OPA sigue siendo Ejecutar |

### 2.3 Escritura 2.4 (misma sesión atómica)

```text
File.Copy plantilla → salida (una vez, igual que HU-07/HU-08/HU-09)
  └─► ValidarFormulasProtegidas (mapa HU-07/HU-08/HU-09 + mapa 2.4: F2:F7, filas 9/11,
      bloque 18–24, CONSOLIDADO J9:J13/K/M, refs DetRetri/DetValiRetri)
        └─► EscribirCeldasLeaf HU-07/HU-08/HU-09 (intactas)
        └─► EscribirCeldasBalanceSc 2.4 (C2:E7 en valores + H si T0-dice-valor)
              └─► revalidar fórmulas protegidas → guardar
```

Ante cualquier fallo: borrar salida parcial (patrón existente). Plantilla origen jamás mutada (hash A4 se mantiene y se extiende al mapa 2.4).

### 2.4 Dominio (Core, sin deps)

```csharp
public sealed class BalanceScAseInputs  // una fila ASE (3–7) de BCE
{
    public Ase Ase { get; set; } = new();
    public decimal Contribucion { get; set; }  // → template-D (hipótesis líder; T0 la prueba)
    public decimal Subsidio { get; set; }      // → template-E (hipótesis líder; T0 la prueba)
    public decimal? Sistema { get; set; }      // → template-H; null = desenlace D2(b) fórmula
    // TotalBSC = Contribucion + Subsidio (propiedad calculada, análoga a ConsolidadoAse.TotalAse)
    public decimal TotalBsc => Contribucion + Subsidio;
}

public sealed class BalanceScInputs  // hoja completa
{
    public IReadOnlyList<BalanceScAseInputs> Ases { get; set; } = ...;  // 5 filas
}
```

`WorkbookLeafInputs` suma `BalanceSc: BalanceScInputs?` (null = comportamiento HU-09 intacto; compatibilidad hacia atrás por construcción). `IWorkbookLeafInputReader`: método de lectura del Balance (contrato intacto, implementación extendida). `ProcesadorPeriodo`: pasos 2.4 integrados al flujo por ASE (leer Balance → leaf → validar gates D5) con fail-fast que nombra ASE; una escritura al final (G5). **Quincena 2:** sin cambios — sigue bloqueada aguas arriba; 2.4 no abre Q2.

> Nota de nomenclatura interna: el modelo usa `Contribucion`/`Subsidio` por significado de dominio, y el **mapa** (`WorkbookLeafCellMapBalanceSc`) fija qué propiedad va a qué columna template según el veredicto T0. Si T0 probara la asignación inversa, solo cambia el mapa, no el modelo.

### 2.5 `IValidador` 2.4 (sin reabrir HU-04..09)

1. Todo lo HU-07/HU-08/HU-09 intacto (matcheo estricto por `Ase.Id`, `AjustesSfT=0` Q1, `GranTotal=Σ`, gates leaf-vs-consolidado y Σ-empresas HU-08, gates banco HU-09).
2. Nuevo (i): por cada ASE: `BCE template == Total General fuente` ±0.5 con la asignación D/E del veredicto T0.
3. Nuevo (ii): por cada ASE: `F == D+E` ±0.5 (aritmética de dominio); sumas fila 11 en dominio ±0.5.
4. Nuevo (iii): H≈F según regla fijada por T0 (redondeo a entero por defecto, a confirmar); vale en ambos desenlaces D2.
5. El validador NO abre `.xlsx` (igual que HU-06..HU-09). J/K/M downstream y bloque 18–24 van a Capa B, no a gates Capa A.

### 2.6 UI — Visual Design Intent (delta mínimo)

Densidad Balanced, mismos GroupBoxes, sin restyle/colores/iconos. El `txtLog` agrega, por cada ASE, una línea BCE (subsidio / contribución / total BSC, "esperado post-Excel", con cita del veredicto D/E) + nota "J9:J13 y K/M calculan por fórmulas (Capa B)". Serilog: mismos eventos HU-07/HU-08/HU-09 con propiedad `Hoja = "BCE SC POR FACT."`. Sin nuevos controles.

### 2.7 Golden Capa A extendida a 2.4 (honestidad HU-06..HU-09)

| # | Qué | Contra qué | Tol |
|---|---|---|---|
| A1 | Celdas BCE escritas en la **salida** (C2:E7 + H si D2(a)) | Mismas celdas **leaf** del golden | ±0.5 |
| A2 | Visibles de **dominio** por ASE (BCE = fuente con asignación T0; F=D+E; H≈F) | Caché golden de filas 3–7 + fila 11 | ±0.5 |
| A3 | F2:F7 + filas 9/11 + bloque 18–24 (+ H si D2(b)) + J/K/M + refs DetRetri/DetValiRetri siguen siendo fórmula en la salida | Estructura | n/a |
| A4 | SHA256 plantilla origen igual antes/después | — | n/a |
| A5 | **Prohibido** comparar caché de fórmula de la salida vs golden; **prohibido** usar CONSOLIDADO J/K/M o DetRetri como oráculo de celdas BCE | — | prohibido |
| A6 | `Conciliaciones/` y `R*_Remuneracion` NO son oráculo | — | n/a |

Capa B (manual Excel: abrir, recalcular, comparar BCE + J9:J13/K9/M9 + fila 9/11 + bloque 18–24 vs golden) fuera de CI, protocolo §5.3.

### 2.8 File changes previstos

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Core/Models/BalanceScInputs.cs` | Crear | `BalanceScAseInputs` + `BalanceScInputs` (§2.4) |
| `Remuneracion.Core/Models/WorkbookLeafInputs.cs` | Modificar | Sumar `BalanceSc` (nullable, default null) |
| `Remuneracion.Core/Interfaces/IWorkbookLeafInputReader.cs` | Modificar | Método lectura del Balance (overload) |
| `Remuneracion.Core/Interfaces/IValidador.cs` | Modificar | Overload/gates D5 (BCE=fuente, F=D+E, H≈F) |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Modificar | Gates D5 |
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` | Modificar | Pasos 2.4 por ASE, fail-fast nombra ASE, una escritura |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapBalanceSc.cs` | Crear | Mapa explícito ASE → editables (congelado T0, incluye veredicto D/E y H) |
| `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` | Modificar | Lectura "Total General" desde el final, match normalizado (D1) |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs` | Modificar | Coherencia BCE=fuente + F=D+E |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Escribir celdas 2.4 en la misma pasada + mapa protegido ampliado |
| `Remuneracion.Infrastructure/FileSystem/ArchivoFuenteLocator.cs` | Modificar | Segundo prefijo `R4-BalanceSubsidioyContribuciones-Optimizado_` (V8) |
| `Remuneracion.WinForms/Form1.cs` | Modificar | Resumen BCE en log |
| `Remuneracion.IntegrationTests/GoldenBalanceScTests.cs` | Crear | Capa A 2.4 (§2.7) |
| `Remuneracion.IntegrationTests/BalanceScTests.cs` | Crear | Dominio: asignación D/E (ambos desenlaces), F=D+E, H redondeo |
| `Remuneracion.IntegrationTests/ProcesadorPeriodoTests.cs` | Modificar | Casos 2.4 (integración período con Balance) |

**No tocar (salvo bug blocker):** `IPlantillaWriter`/validation-only (HU-04); `IRecaudoReader` agregados HU-02; coherencia `F25`-Extemp HU-05; semántica single-ASE y multi-ASE HU-07; mapa por empresa HU-08; mapa banco HU-09; `DetRetriRounder` (2.6); `requirements/` legado.

---

## 3. SPEC / Acceptance Criteria (trazabilidad Propuesta §10 + base docs)

### Requirement 1 — Lectura del "Total General" (CA-1; Detalle de plantilla Inst. BCE; Proceso de Recaudo § balance)

El sistema **MUST** leer la fila "Total General" del final de cada `R4-BalanceSubsidioyContribuciones` (búsqueda dinámica desde el final, match normalizado — tolera variantes de mayúsculas/espacios). **MUST NOT** usar fila fija; si falta la etiqueta → fallo que nombra ASE, nunca valor inventado. La variante Optimizado de ASE5 **MUST** resolverse por el segundo prefijo del locator.

- GIVEN carpeta `1-Promoambiental` Q1 → WHEN lectura Balance → THEN E-fuente/Subsidio y F-fuente/Contribución tales que, con la asignación T0, D-template=`3256235169.97` y E-template=`-1223871491.13` ±0.5.
- GIVEN fuente ASE5 Optimizado → THEN misma lectura por el prefijo variante, con paridad de columnas probada por T0.

### Requirement 2 — Asignación D/E probada, no asumida (CA-2)

El sistema **MUST** aplicar la asignación D/E del veredicto T0 (§0.2): hipótesis líder D←F-fuente/E←E-fuente salvo prueba en contrario registrada. **MUST NOT** codificar la asignación del texto del doc ni la de los headers sin la prueba fuente-vs-template.

- GIVEN veredicto T0 hipótesis-líder → THEN `BalanceScAseInputs.Contribucion` → template-D y `Subsidio` → template-E.
- GIVEN veredicto T0 inverso → THEN solo cambia `WorkbookLeafCellMapBalanceSc` (el modelo no cambia) + nota de divergencia headers-vs-texto.
- GIVEN T0 sin match ±0.5 → THEN HU recortada y escalada; **MUST NOT** inventar mapeo.

### Requirement 3 — Coherencia BCE (CA-2)

El sistema **MUST** cumplir (i) BCE por ASE = Total General fuente ±0.5 y (ii) F=D+E ±0.5 por ASE. **MUST NOT** presentar valores CONSOLIDADO J/K/M o DetRetri como valores BCE.

- GIVEN leafs 2.4 Q1 → THEN fila ASE1 §2.1 exacta ±0.5 + equivalentes ASE2..ASE5 de T0-0.5.

### Requirement 4 — Escritura 2.4 preservando fórmulas (CA-3/CA-4; Prompt Maestro Vo paso valores)

El sistema **MUST** escribir solo celdas del mapa BalanceSc (C2:E7 + H si D2(a)) en la misma pasada HU-07/HU-08/HU-09; **MUST NOT** escribir ninguna celda con `<f>` (incluye F2:F7, filas 9/11, bloque 18–24, J/K/M, refs DetRetri/DetValiRetri, resto 2.5–2.7); **MUST NOT** mutar la plantilla; ante fallo **MUST** borrar la salida parcial.

- GIVEN 5 leafs Balance válidos → WHEN overload lista → THEN cambian solo celdas del mapa; A3 verde.
- GIVEN mismatch en ASE4 → THEN `CalculoInvalidoException` que nombra ASE, sin archivo certificado.

### Requirement 5 — Columna H según T0 (CA-3)

El sistema **MUST** tratar H según el veredicto D2: (a) valor → escribir con la regla de redondeo fijada por T0 (por defecto ROUND(F,0)); (b) fórmula → proteger y verificar H≈F. **MUST NOT** asumir un desenlace antes de T0.

- GIVEN ASE1 F=`2032363678.84` → THEN H=`2032363679` ±0.5 en ambos desenlaces.

### Requirement 6 — Trazabilidad y UI honesta (CA-6/CA-7)

Serilog + `txtLog` por ASE (qué fuente→qué celdas, BCE vs fuente, F=D+E, H≈F, veredicto D/E citado). Resumen etiqueta "esperado post-Excel". Sin restyle.

### Requirement 7 — Golden Capa A 2.4 (CA-3/CA-5)

Matriz §2.7 para 5 ASE con insumos Q1. **MUST NOT** comparar caché de salida vs golden (A5).

| CA §10 | HU-10 |
|---|---|
| CA-1 | Lee el Total General de los 5 ASE (fail-fast nombra ASE; Optimizado por segundo prefijo) |
| CA-2 | BCE = fuente (asignación T0) y F=D+E (agregados J/K/M ≠ valores BCE) |
| CA-3 | Capa A 2.4; J/K/M + 18–24 correctos post-Excel (Capa B manual residual) |
| CA-4 | Reassert mapa ampliado + plantilla no mutada (hash) |
| CA-5 | Gates D5 + coherencia leaf Balance |
| CA-6 | Serilog + log BCE por ASE |
| CA-7 | Mismo flujo 5-ASE + resumen BCE |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 500–900 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 T0 + mapa (+veredictos D/E y H) → PR2 lectura Balance + locator 2 prefijos → PR3 writer + mapa protegido → PR4 validador + coherencia + período/UI → PR5 golden + tests |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes. Chained PRs recommended: Yes. Chain strategy: pending.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 0 | Discovery T0 + mapa congelado (+D/E y H) | PR 1 | Bloquea todo; solo lectura + datos |
| 1 | Lectura Balance + locator | PR 2 | Depende de PR 1 |
| 2 | Writer 2.4 + protegido ampliado | PR 3 | Depende de PR 1 |
| 3 | Validador D5 + período + UI | PR 4 | Depende de PR 1 |
| 4 | Golden 2.4 + tests + Capa B doc | PR 5 | Depende de PR 2–4; 47 tests previos verdes |

### Phase 0 — Discovery T0 (BLOQUEANTE, solo lectura)

- [ ] 0.1 Clasificar valor-vs-fórmula celda por celda: BCE C2:E7, F2:F7, H2:H7 (¿H es valor o `=ROUND(F,0)`?), I2:I7, filas 9/11, bloque 18–24. Registrar fórmula visible de cada celda con `<f>`.
- [ ] 0.2 **Resolver la contradicción D/E** (§0.2): leer la fila "Total General" de la fuente ASE1 (E-fuente, F-fuente en valores) y matchear contra D3/E3 del golden ±0.5; repetir para ASE2..ASE5; emitir veredicto congelado (hipótesis líder / inverso / fallido).
- [ ] 0.3 Localizar el "Total General" en las fuentes ASE2..ASE5 (etiqueta exacta, distancia al final, variantes) y verificar paridad de layout de la variante Optimizado ASE5 (columnas B..G, hoja, etiqueta).
- [ ] 0.4 Fijar la normalización de etiqueta ("total general" lowercase/trim) y la regla H (redondeo exacto: ¿ROUND(F,0)? ¿otra?) + semántica de columna I (¿fórmula de diferencia?).
- [ ] 0.5 Extraer del caché golden las filas 3–7 de ASE2..ASE5 (equivalente a §2.1) + verificar 2.5–2.7/CONSOLIDADO-JKM/DetRetri como fórmulas + barrido de externalLinks en BCE.
- [ ] 0.6 Determinar el veredicto D2(a)/D2(b) para columna H y el tratamiento de C (¿valor ASE id o etiqueta?).
- [ ] 0.7 Congelar `WorkbookLeafCellMapBalanceSc` (ASE → editables C/D/E(+H) + visibles + veredicto D/E + veredicto H). **Nada entra al código sin esta tabla.**

### Phase 1 — Dominio (modelos + lectura)

- [ ] 1.1 `BalanceScInputs.cs` (§2.4; TotalBsc como propiedad calculada) + extender `WorkbookLeafInputs` (nullable, default null).
- [ ] 1.2 Reader Balance según T0-0.7 (búsqueda desde el final, match normalizado, asignación D/E del veredicto); fallo nombra ASE.
- [ ] 1.3 `ArchivoFuenteLocator`: segundo prefijo `R4-BalanceSubsidioyContribuciones-Optimizado_` (V8; aceptar ambos, base primero).

### Phase 2 — Escritura (misma pasada)

- [ ] 2.1 `WorkbookLeafCellMapBalanceSc.cs` (datos T0-0.7) + escritura de celdas 2.4 dentro del overload lista existente.
- [ ] 2.2 Mapa protegido ampliado (F2:F7, filas 9/11, bloque 18–24, J/K/M, refs DetRetri/DetValiRetri, resto 2.5–2.7; shared-formula awareness HU-07T0 §0.5).
- [ ] 2.3 Borrado de parcial + hash plantilla intactos.

### Phase 3 — Orquestación + UI delta mínimo

- [ ] 3.1 `ProcesadorPeriodo`: pasos 2.4 por ASE (leer Balance → leaf → validar D5), fail-fast nombra ASE, una escritura.
- [ ] 3.2 `Form1`: resumen BCE por ASE en log (§2.6).
- [ ] 3.3 Serilog BCE por ASE (propiedad `Hoja`, mismos sinks).

### Phase 4 — Pruebas y evidencia

- [ ] 4.1 `BalanceScTests` (in-memory + fuentes Q1 reales sin salida): asignación D/E en ambos desenlaces (test parametrizado líder/inverso), F=D+E ok, H redondeo ok/ko, label variant "TOTAL GENERAL" matchea, mismatch nombra ASE, J/K/M usados como valor BCE (prohibido — el test lo demuestra fallando).
- [ ] 4.2 `GoldenBalanceScTests`: matriz §2.7 (5 ASE, fixtures Q1).
- [ ] 4.3 Integración período con Balance + negativa (falta "Total General" en fuente ASE3 → ASE en el error, sin salida; variante Optimizado ASE5 resuelve por segundo prefijo).
- [ ] 4.4 Los 47 tests existentes verdes; build 0 warnings; CRLF; sin commit.

### Phase 5 — Documental

- [ ] 5.1 Capa B manual §5.3 ejecutada una vez y evidenciada (sin fingirla como gate de merge).
- [ ] 5.2 Cierre deja explícito el frente 2.5 (siguiente HU propuesta: AJUSTES-SF-T, quincena 2).

---

## 5. Test Strategy

| Capa | Qué | Cómo |
|---|---|---|
| Unidad | Gates D5 (BCE=fuente con asignación T0, F=D+E, H≈F) | In-memory, fila §2.1 + T0-0.5 |
| Unidad | Reader Balance (búsqueda final, label variants, Optimizado ASE5) | Fuentes Q1 reales, sin Excel de salida |
| Unidad | Asignación D/E parametrizada (líder e inverso) | In-memory; el mapa decide, el modelo no cambia |
| Integración | Período con 2.4 (5 carpetas reales, salida temp) | Insumos Q1, fail-fast nombra ASE |
| Golden Capa A 2.4 | Matriz §2.7 | OpenXML read-only + aritmética dominio; plantilla=golden (A5 aplica) |
| UI | Resumen BCE | Funcional manual (sin harness) |
| Capa B | BCE + J/K/M + 9/11 + 18–24 post-Excel | Manual — §5.3 |

### 5.1 Fixtures

- Golden/plantilla: `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` (hoja `BCE SC POR FACT.` = sheet9).
- Fuentes: `Docs/Insumos/REMUNERACION 2026071/{1..5}-*/R4-BalanceSubsidioyContribuciones*.xlsx` (Q1; incluye variante `-Optimizado_` en ASE5; Q2 `REMUNERACION 2026072/` bloqueada — sigue lanzando `CalculoInvalidoException` aguas arriba).
- Referencia: fila §2.1 (ASE1) + T0-0.5 (ASE2..ASE5), tolerancia ±0.5.

### 5.2 Casos negativos obligatorios (nombran ASE)

"Total General" sin etiqueta final en ASE3; variante "TOTAL GENERAL" con mayúsculas debe matchear (no fallar); variante Optimizado pedida por prefijo base en ASE5 (debe resolver igual); asignación D/E invertida a propósito (el test parametrizado lo detecta); Σ F ≠ D+E en ASE2; H ≠ ROUND(F,0) fuera de ±0.5; salida == plantilla (no in-place); overwrite cancelado; Q2 sigue bloqueada; J9 CONSOLIDADO usado como valor BCE (prohibido — el test lo demuestra fallando); escritura en DetRetri con `-BCE!F3` (prohibida — el mapa protegido la rechaza).

### 5.3 Protocolo manual Capa B (no CI)

1. Generar salida a ruta distinta del golden. 2. Abrir en Excel, recalcular. 3. Comparar BCE C2:E7/H/F + fila 9/11 + bloque 18–24 vs golden ±0.5 (refs §2.1 + T0-0.5); verificar J9:J13 = F3:F7, K9 = D9-E9-F9-G9-H9-I9-J9, M9 = K9-L9; verificar DetRetri/DetValiRetri recalcularon con `-BCE!F3..F7`. 4. Evidencia adjunta. No es gate de merge.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | Balance = extensión de datos (mapa + inputs + reader + gates); período/orquestación/writer conservan su rol HU-07/HU-08/HU-09. |
| **O** | Se agregan modelos/mapas/overloads; el path HU-09 funciona con `BalanceSc = null` (abierto sin modificar). |
| **L** | `OpenXmlPlantillaWriter` escribe más celdas del mismo modo leaf; comportamiento HU-09 puro inalterado. |
| **I** | `BalanceScInputs` separado de `ReporteBancoInputs`/`ConciliacionEmpresaInputs`; reader/validador crecen por overload. |
| **D** | Core define Balance/inputs/gates; Infrastructure/WinForms componen. Sin nuevas deps. |

### 6.2 Best Practices

- La verdad del workbook manda: F2:F7 + 9/11 + 18–24 + J/K/M + refs DetRetri/DetValiRetri ⇒ cero escritura directa (Rector §6 preservación).
- La contradicción D/E se resuelve con prueba fuente-vs-template, no con autoridad del texto ni de los headers (honestidad técnica §0.2).
- Mapa explícito por ASE verificado, no offsets ni filas fijas (Rector §11.3; V6/V8).
- Una escritura atómica; plantilla nunca mutada; hash A4 extendido.
- Golden honesto BCE (OpenXML no recalcula; A5/A6).
- Fail-fast nombra ASE; sin salida certificada ante fallo.
- Documentos base citados por sección como spec funcional (§1.5), con la divergencia D/E declarada en vez de ocultada.
- Tándem HU-09: mismo patrón de extensión leaf; HU-10 no duplica ni anticipa 2.5–2.7.

### 6.3 Performance

- 5 filas × ~4 celdas ≈ <25 escrituras extra en la misma sesión OpenXML + 5 lecturas de "Total General" (búsqueda desde el final sobre reportes por localidad, ExcelDataReader streaming). Irrelevante a esta escala; `Task.Run` existente para no congelar el form.

**Veredicto:** APROBADO como it. 2.4 de Fase 2 **si** T0 congela el mapa con evidencia (incluidos los veredictos D/E y H) y se acepta CA-3 parcial (Capa A en CI, Capa B manual).

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| La prueba D/E no cierra ±0.5 en ningún sentido (fuente y template no matchean) | Baja | T0-0.2 declara veredicto fallido; HU recortada y escalada al Ingeniero; prohibido inventar mapeo |
| El veredicto D/E sale inverso a los headers (el texto del doc tenía razón) | Baja | D4: solo cambia el mapa, no el modelo; se documenta la divergencia headers-vs-texto |
| Columna H es fórmula con semántica distinta a ROUND(F,0) | Media | D2 doble desenlace + T0-0.4; el gate H≈F usa la regla fijada por T0 |
| El "Total General" de ASE2..ASE5 trae etiqueta o layout distintos (incluido Optimizado) | Alta | D1 (búsqueda + normalización) + T0-0.3; si el hueco es real, la HU se recorta a los ASE verificados con rebase, no se inventa |
| Alguna fila ASE falta cuando el neto es 0 | Baja | T0-0.5; el reader distingue "fila ausente = 0 demostrable" de "fuente ausente = fallo" (regla fijada en T0-0.7) |
| Inflar a 2.5–2.7 dentro de esta HU (DetRetri tienta, 18–24 tienta) | Media | §0.3 out-of-scope + D6 (protegidas, no implementadas); rechazar PRs que lo metan |
| Comparar caché de salida vs golden y "cerrar" CA-3 en falso | Alta | A5 en tests; este plan lo prohíbe |
| Carpeta `5-Área Limpia` (tilde) + variante `-Optimizado_` en nuevos tests | Media | Reutilizar `Insumos.cs` (ya maneja la tilde); el locator prueba ambos prefijos |
| Q2 (`REMUNERACION 2026072/`) confundida con fixture Q1 | Media | Fixtures §5.1 fijan Q1; Q2 sigue bloqueada aguas arriba |

---

## 8. Rollback

- Revertir PRs en orden inverso (golden → período/UI → validador → writer → reader/locator → mapa/T0).
- HU-04..HU-09 intactas sin esta HU: `BalanceSc = null` = comportamiento HU-09 puro.
- `Docs/Insumos/` untracked: jamás como destino; pisada se restaura desde backup.
- Si T0 demuestra hueco de fuente o veredicto D/E fallido (Riesgos 1–2), recorte con rebase, no invención.

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §9 it. 2.4 / §6 / §10 (±0.5) + tres documentos base como origen funcional (§1.5). 2.5–2.7, INTERVENTORIA y Q2 quedan fuera, citados a §9.
2. T0 bloqueante con protocolo de resolución D/E (§0.2): hipótesis líder D←F-fuente/E←E-fuente a probar; sin prueba no hay código.
3. Lectura del "Total General" por búsqueda dinámica desde el final con match normalizado (V6); fila fija y string exacto prohibidos; locator con dos prefijos (V8).
4. F2:F7 + filas 9/11 + bloque 18–24 + J/K/M + refs DetRetri/DetValiRetri = fórmulas protegidas (hipótesis V3/V4/V5, T0 confirma); su lógica es 2.6/2.7, no 2.4.
5. Columna H con doble desenlace T0 (D2): valor → se escribe con redondeo T0; fórmula → se protege y solo se verifica H≈F.
6. Una sola escritura atómica (overload lista HU-07/HU-08/HU-09); sin segundo pase.
7. Gates D5 (BCE=fuente con asignación T0, F=D+E, H≈F) ±0.5; fail-fast nombra ASE.
8. Mapa explícito por ASE congelado por T0; ninguna dirección sin evidencia.
9. Golden Capa A 2.4 con honestidad HU-06..HU-09; Capa B manual residual.
10. UI delta mínimo + Serilog BCE por ASE; OPA = Ejecutar.
11. Apply espera aprobación + PRs encadenados (Unidad 0–4, §4).

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
