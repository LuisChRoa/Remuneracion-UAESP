# Plan 08 — HU-08: Conciliación por empresa de facturación (Fase 2, it. 2.2)

> **Historia:** diligenciar las hojas de conciliación por empresa de facturación (EAAB Reciprocidad, ENEL, ENERBIT, OCCIDENTE Directa, EAAB+CiudLimp) para los 5 ASE, reutilizando la ruta leaf multi-ASE certificada de HU-07.
> **Rector (irrenunciable):** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` — §9 Fase 2 it. **2.2 Hojas de conciliación**, §6 arquitectura (capas, preservación de fórmulas, Serilog, lectura directa de fuente), §10 CA + tolerancia ±0.5. **Entra SOLO 2.2.** Salen 2.3 REPORTE x BANCO, 2.4 BCE, 2.5 AJUSTES-SF-T, 2.6 DetRetri, 2.7 validaciones cruzadas — son HUs posteriores. Quincena 2 sigue bloqueada.
> **Continuidad:** HU-01..HU-07 cerradas (ruta leaf multi-ASE, mapa explícito por `Ase.Id`, fórmulas intactas, build 0/0, tests 25/25, harness 24/24). Este plan NO reabre su semántica.
> **Estado:** PENDIENTE DE APROBACIÓN — no implementar hasta OK del Ingeniero.
> **Fecha:** 2026-09-08

---

## 0. Clarification Gate

No hay pregunta bloqueante para el Ingeniero. El descubrimiento previo a este plan resolvió la arquitectura (las hojas `REMUNERACION_*` son 100 % derivadas); lo que queda abierto es el mapa celda-por-celda de las filas detalle por empresa, y por eso el plan incluye discovery T0 obligatorio (§4 Fase 0). El apply espera aprobación explícita.

### 0.1 Qué está verificado y qué NO (honestidad técnica)

**Verificado en esta planificación** (dump OpenXML raw de `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` + `mcp-excel` + `unzip` de los `.docx`):

| # | Hecho | Método | Consecuencia |
|---|---|---|---|
| V1 | El libro tiene **40 hojas**. Las 5 de conciliación son `REMUNERACION Reciprocidad EAAB`, `REMUNERACION_ENEL`, `REMUNERACION_ENERBIT`, `REMUNERACION_OCCIDENTE_ Directa`, `REMUNERACION_EAAB-CL`, cada una con su `VALIDACION_*` y `GERENTES_*` | `list_sheets` | Alcance de hojas 2.2 cerrado: 5 + 5 + 5 (las `VALIDACION_*`/`GERENTES_*` solo se protegen, no se escriben — ver V4) |
| V2 | `REMUNERACION_ENEL` = bloques apilados **R1 Oportuno / R2 Anticipos / R3 Extemporáneo / R4 Reversión / R5 Saldos-por-nota (ceros Q1) + Reporte consolidado // ENEL + Reporte consolidado de anticipos + Validación del recaudo vs remuneración**. Cada bloque trae sus 5 ASE + fila TOTAL | Lectura `A1:Z140` | Layout por empresa conocido a nivel bloque; direcciones exactas por bloque las fija T0 |
| V3 | Las 5 hojas `REMUNERACION_*` son **100 % fórmulas hacia `Reporte Componentes R1` / `Rem. Anticipos R2` / `Reversion Pagos R4` / `CONSOLIDADO_TOTAL RECAUDO`**. Conteo `<f>`: RECIP 1007 (R1 150, R2 75, R4 75), ENEL 1039 (150/75/75), ENERBIT 1029 (150/75/75), OCCIDENTE 969 (150/75/75), EAAB-CL 964 (**R1 145, R2 70, R4 70**). Ej. ENEL `D9='Reporte Componentes R1'!F56` + caché `16180195504.29` | Dump `<f>` sheet21/24/27/30/33 | **Decisión arquitectónica central:** en `REMUNERACION_*` NO hay nada que escribir. La superficie escribible de 2.2 vive en R1/R2/R4 (filas detalle por empresa) + hojas `Recaudo *` |
| V4 | `VALIDACION_ENEL` compara `Recaudo ENEL` vs `REMUNERACION_ENEL` (`O3=H3-N3`, `P3=INT(O3)=0` — diferencia debe ser 0). Es contenido de **2.7**, no de 2.2 | Dump `<f>` sheet25 | `VALIDACION_*`/`GERENTES_*` entran al mapa de fórmulas protegidas; su lógica NO se implementa en esta HU |
| V5 | `Reporte Componentes R1` trae **filas detalle por empresa dentro de cada bloque ASE**: ASE1 r51 RECIPROCIDAD / r56 ENEL / r61 ENERBIT / r66 OCCIDENTE / r71 CIUDAD LIMPIA-ACUEDUCTO; ASE2 r181–201; ASE3 r321–341; ASE4 r442–462; ASE5 r524–544. Las filas con recaudo son **fórmulas** (`F56=F33+F14-L14`; `G56:AP56` shared `=G14`); las de valor 0 son **valores estáticos** (`F51/F61/F71 = 0` sin `<f>`) | Dump `<c>` raw sheet10 + shared-formula scan (845 shared de 915 `<f>`) | El writer de 2.2 escribe **operandos** de esas fórmulas (filas `Total` por empresa de la zona principal, ej. r33/r14), jamás las filas detalle. Direcciones por bloque las congela T0 |
| V6 | `Rem. Anticipos R2` trae filas por empresa en **col B** (ASE1 r44/47/50/53/56; ASE2 r138–150; ASE3 r250–262; ASE4 r346–358; ASE5 r416–428). `Reversion Pagos R4` en **col A** (ASE1 r70–82; ASE2 r164–176; ASE3 r201–213; ASE4 r315–327; ASE5 r350–362) | Scan shared-strings sheet11/sheet12 | Superficie R2/R4 confirmada a nivel filas; carácter valor-vs-fórmula por celda lo fija T0 |
| V7 | **Coherencia Σ empresas = TOT_OPT verificada 5/5** con caché golden: ASE1 `0+16180195504.29+0+524136930.28+0=16704332434.57` ✓ D9; ASE2 `13570060+11762458070.29+0+381752310.90+0=12157780441.19` ✓ D10; ASE3 `0+9648710791.68+0+453277723.14+0=10101988514.82` ✓ D11; ASE4 `7264980+10318956340.32+0+226188183.51+0=10552409503.83` ✓ D12; ASE5 `0+8229736739.09+126372775.04+163200815.15+0=8519310329.28` ✓ D13 | Aritmética sobre caché OpenXML del golden | Es el **gate de coherencia central** de HU-08 (análogo al gate HU-05 `F25` vs Extemp). Ceros por empresa son legítimos (matriz dispersa, ver V8) |
| V8 | Matriz empresa×ASE Q1 real (caché bloque R1, fila TOTAL incluida): RECIP `[0, 13570060, 0, 7264980, 0]` Σ 20835040; ENEL 5/5 con valor Σ 56140057445.67; ENERBIT solo ASE5 (126372775.04); OCCIDENTE 5/5 Σ 1748555962.98; EAAB-CL todo 0 | Caché `D9:D14` sheet21/24/27/30/33 | Compatible con Apéndice B del Rector (ASE1 ENEL+Occidente; ASE2 ENEL+EAAB Recip; ASE3 ENEL+Occidente; ASE4 ENEL+EAAB Recip; ASE5 ENEL+ENERBIT+Occidente). EAAB-CL sin recaudo en Q1: el gate debe **permitir ceros**, no exigir no-cero |
| V9 | Las hojas `Recaudo *` son **inputs en valores**: `Recaudo ENEL!D9 = 55775447693` sin fórmula; solo 15 `<f>` (totales/diferencias). Docx: *"En esta hoja se pega todo en valores. De la fila 29 en adelante, validaciones de sumas"* por hoja y empresa | Dump sheet2 + `Detalle de plantilla.docx` (legible vía unzip, 26204 chars) | Las 5 hojas `Recaudo *` SÍ se escriben (zona de datos filas ~3–28; fila 29+ protegida). Fuente de sus valores la fija T0 (candidatos §2.2) |
| V10 | La **fuente R1 ya trae desglose por empresa**: `Recaudoporcomponente` col C = `ENEL`/`OCCIDENTE` + col D `Total`, cols F–L por componente (TDF/TTL/TVIAT/Aprovechamiento/CCSA/Especiales). Ej. ASE1 ENEL `F=19104244534.3` | Lectura `Sheet1 A1:L45` fuente ASE1 | Hipótesis primaria: R1-detalle ← misma fuente R1 (doc/string del docx: *"validación por fuente de recaudo"*). T0 la confirma fila-por-fila |
| V11 | `RecaudosReversados` trae detalle por banco/canal (OCCIDENTE, T23-OCCIDENTE OFC…) + bloque `66/ASEO`: `1-APLICADOS A FACTURACION 11673020`, `2-SALDOS A FAVOR 1944280`, `9-REVERSION PAGOS -12054255.65`, `9-REVERSION SALDOS -8965804.35`. Docx: *"De la fila 9 a la 58, el recaudo recibido para cada ASE: 1-APLICADOS…, 2-SALDOS…"* | Lectura fuente ASE1 + docx | Hipótesis primaria para hojas `Recaudo *`. T0 confirma mapeo concepto→celda |
| V12 | `Consolidado/Conciliaciones/Conjunta *-072026*.xlsx` (una hoja `RESUMEN MES`): OPORTUNO / EXTEMP / TOTAL por ASE con VALOR + N° REG, Q2 = 0. Sus valores **≠** `REMUNERACION_*` (ej. ASE1 oportuno ENEL 16020970408 vs 16180195504.29) | Lectura `RESUMEN MES A1:L30` | Son **control**, no copia directa. Dirección insumo→plantilla la fija T0; prohibido asumirla |
| V13 | `Consolidado/Remuneración del sistema/R{1..4}_Remuneracion_2026071.xlsx` = hojas `DetRetri/DetValiRetri` (sistema PROCERASEO). Contexto de **2.6**, no insumo de 2.2 salvo que T0 los use como control | `list_sheets` | Fuera de 2.2; se citan para no confundirlos con insumos |
| V14 | `.docx` **legibles** (citas §1.5). `Proceso de Recaudo.docx` (7359 chars): las EFC reportan TXT diarios por banco/canal (no disponibles como insumos) — por eso 2.2 se alimenta de los reportes por ASE, no de TXT | Unzip + extracción | Cierra la puerta a "leer los TXT de las EFC": no existen en insumos; no son alcance |
| V15 | Contratos listos para extender: `IWorkbookLeafWriter` ya tiene overload lista multi-ASE; `WorkbookLeafInputs{R1,R2,R4}` usan `CeldasPorAse` (dict ref→valor); `Insumos.cs` + `CarpetasAse.Prefijos` cubren las 5 carpetas | Lectura de código | 2.2 = **extensión del patrón HU-07**, no arquitectura nueva |

**NO verificado (y por eso T0 es bloqueante, §4 Fase 0):**

1. Direcciones exactas de los **operandos** por empresa y bloque en R1 (zona principal r~9–40 por bloque: ¿qué fila `Total` alimenta a cada `F{detalle}`? — ej. ¿`F33/F14/L14` vienen de fuente y con qué labels?), R2 (operandos de cada fila empresa col B) y R4 (idem col A). Solo se dumpeó ASE1-R1-parcial.
2. Carácter valor-vs-fórmula de cada operando candidato (el caso `F56=F33+F14-L14` demuestra que hay que clasificar antes de prescribir escritura).
3. Mapeo fuente→template para R2/R4 por empresa: ¿la fuente R2/R4 trae desglose por empresa (qué labels)?; si no, ¿`RecaudosReversados` / `ReportePagosxBanco` / `Balance` lo aportan, o el proceso manual lo deriva?
4. Mapeo fuente→celda de las 5 hojas `Recaudo *` (filas ~3–28): ¿qué filas/columnas de `RecaudosReversados` alimentan cada concepto por ASE? ¿`Conciliaciones` entra como input o solo control?
5. Por qué EAAB-CL tiene 5 `<f>` menos por bloque (145/70/70 vs 150/75/75): ¿bloque R5 ausente o fila menos? (compatible con R5-ceros, pero T0 lo confirma).
6. Si `GERENTES_*` son fórmulas puras (esperado) o traen valores pegados: T0 lo verifica antes de meterlas al mapa protegido.
7. Exhaustividad de links externos en las 5 hojas (muestra de sheet24: solo refs internas; externalLinks del libro apuntan a tarifas/bancos de otras hojas — T0 verifica que ninguna celda objetivo de 2.2 dependa de un externalLink roto).

> **Regla de hierro del plan:** ninguna dirección de celda del mapa por-empresa entra al código sin pasar por T0. Lo ya verificado arriba (V1–V15) sí es contratable desde el día uno.

### 0.2 Mapeo al Rector (in vs out)

**Entra porque §9 it. 2.2 / §6 / §10 lo piden ahora:**

| Rector | Qué cubre HU-08 |
|---|---|
| §9 it. 2.2 | Diligenciar hojas por empresa de facturación (EAAB Reciprocidad, ENEL, ENERBIT, OCCIDENTE Directa, EAAB+CiudLimp) |
| §7.1 bloque escritura | "Recaudo EAAB/ENEL/etc → pega conciliación por empresa" + detalle R1/R2/R4 por empresa que alimenta `REMUNERACION_*` |
| §10 CA-1/CA-2 | Leer las fuentes por empresa de los 5 ASE + cálculo/coherencia por empresa (Σ empresas = visible de bloque, §2.5) |
| §10 CA-4/CA-5/CA-6/CA-7 | Fórmulas intactas (mapa ampliado 2.2), validación por empresa, Serilog por empresa, UI sin cambios salvo resumen |

**Sale porque §9 lo asigna a 2.3–2.7 / Fase 3 (HUs posteriores):**

- 2.3 `REPORTE RECAUDO x BANCO` (filas 9–58, E59:E80 diferencias por anulados — docx).
- 2.4 `BCE SC POR FACT.` (Balance Subsidio/Contribuciones).
- 2.5 AJUSTES-SF-T (`SALDOS POR NOTA` + `RETRIBUCION NEGATIVA`; Q1 sigue con `AjustesSfT=0`; quincena 2 sigue bloqueada en `CalculoRemuneracion`).
- 2.6 `DetRetri2026071` / `DetValiRetri2026071` (enteros redondeados como objetivo de escritura).
- 2.7 validaciones cruzadas (`VALIDACION_TOTAL`, `VALIDACION_RECIP/ENEL/…`, `Valida -*`). Se protegen sus fórmulas; su lógica no se implementa.
- `INTERVENTORIA`, `ANT EXT-REV`, `ANTICIPOS USUARIOS`, `Informe AFaseo Recaudo` (se ignora para el consolidado, docx Inst. 9).

### 0.3 Decisiones ejecutivas (aprobación = aceptar estas)

| ID | Decisión |
|---|---|
| G1 | **Ninguna escritura directa en `REMUNERACION_*`**: son 100 % fórmulas (V3). Se validan como protegidas; el trabajo 2.2 vive en R1/R2/R4 + `Recaudo *`. |
| G2 | **Una sola llamada de escritura** por proceso (mismo overload lista HU-07). Sin segundo pase: la superficie 2.2 cabe en la misma sesión `File.Copy` + validación pre/post ampliada. Ver §2.3. |
| G3 | **Mapa por empresa explícito** (`WorkbookLeafCellMapPorEmpresa`: empresa × ASE → refs editables), no offsets. V5/V6 muestran layouts heterogéneos por hoja (col C vs B vs A) y EAAB-CL con 5 fórmulas menos. |
| G4 | **Discovery T0 bloquea el mapa por empresa** (§4 Fase 0). Sin dump operando-por-operando no hay código 2.2. |
| G5 | **Gate de coherencia Σ empresas = visible de bloque** por ASE y hoja (R1/R2/R4), tolerancia ±0.5, con ceros legítimos (V7/V8). Es el equivalente 2.2 del gate HU-05. |
| G6 | **`Recaudo *` se escribe en valores** (V9, docx), filas ~3–28; fila 29+ validaciones protegidas. Fuente primaria hipotética: `RecaudosReversados`; `Conciliaciones` = control (V12). T0 cierra la dirección. |
| G7 | **UI sin cambios funcionales**: el modo 5-ASE ya existe; 2.2 solo agrega líneas de resumen por empresa al log. Sin restyle. |
| G8 | **Golden Capa A extendida a 2.2** con honestidad HU-06/HU-07 (leaf salida vs leaf golden; visibles de dominio vs caché golden; nunca caché de salida vs golden). |

---

## 1. PROPOSE

### 1.1 Intent

Diligenciar la conciliación por empresa de facturación para los 5 ASE dentro de la misma escritura atómica del período: filas detalle por empresa en R1/R2/R4 + hojas `Recaudo *` en valores, de modo que las 5 hojas `REMUNERACION_*` calculen solas por fórmulas y la coherencia Σ empresas = visible de bloque quede demostrada contra el golden Q1.

### 1.2 In Scope

- Discovery T0 (§4 Fase 0) + `WorkbookLeafCellMapPorEmpresa` congelado solo con evidencia.
- Lectura por empresa desde fuentes por ASE (mismos 6 xlsx por carpeta; sin nuevos tipos de archivo salvo que T0 demuestre que `Conciliaciones/` es input — en cuyo caso se registra como hallazgo y se recorta o rebasea, no se improvisa).
- Extensión de `WorkbookLeafInputs{R1,R2,R4}` + `Recaudo*Inputs` con operandos por empresa; escritura en la misma pasada HU-07; validación pre/post con mapa ampliado (incluye `REMUNERACION_*`, `VALIDACION_*`, `GERENTES_*`, `Recaudo *` fila 29+ como protegidas).
- Gate Σ empresas = visible por ASE/hoja en `IValidador` (+ coherencia leaf).
- Golden Capa A extendida a 2.2 + tests con insumos Q1 reales.
- Resumen por empresa en log/Serilog (delta mínimo UI).

### 1.3 Out of Scope

Todo §0.2 (2.3–2.7, Fase 3). Además: quincena 2 (sigue lanzando `CalculoInvalidoException`); reescritura del path HU-07 (se extiende, no se duplica); leer TXT diarios de EFC (no existen en insumos, V14); restyle UI; DI framework.

### 1.4 Resultado de negocio

El Ingeniero ejecuta el modo 5 ASE como hoy; obtiene el mismo workbook más las filas por empresa y las hojas `Recaudo *` diligenciadas; las hojas `REMUNERACION_*` muestran los valores del golden post-Excel (tabla §2.1); el log audita por ASE **y por empresa**; las pruebas demuestran coherencia Σ empresas = visible por bloque.

---

## 2. DESIGN

### 2.1 Mapa verificado por empresa (contratable desde el día uno)

**Fórmulas `REMUNERACION_*` → detalle R1/R2/R4 (ya son fórmulas; no se tocan):**

| Hoja empresa | R1-block apunta a | R2-block apunta a | R4-block apunta a | Nota |
|---|---|---|---|---|
| `REMUNERACION_ENEL` | `R1!F56 / F186 / …` (ENEL por bloque) | `R2!…` (filas empresa ENEL) | `R4!…` (idem) | `D9=R1!F56` verificado; resto por patrón + T0 |
| `REMUNERACION Reciprocidad EAAB` | `R1!F51 / F181 / …` | idem RECIPROCIDAD | idem | 150/75/75 fórmulas |
| `REMUNERACION_ENERBIT` | `R1!F61 / …` | idem ENERBIT | idem | 150/75/75 |
| `REMUNERACION_OCCIDENTE_ Directa` | `R1!F66 / …` | idem OCCIDENTE | idem | 150/75/75 |
| `REMUNERACION_EAAB-CL` | `R1!F71 / …` (145, no 150) | idem (70, no 75) | idem (70, no 75) | T0-0.5 explica la diferencia |

**Filas detalle por empresa en R1/R2/R4 (filas verificadas; columnas/carácter los fija T0):**

| Bloque ASE | R1 (col C) | R2 (col B) | R4 (col A) |
|---|---|---|---|
| ASE1 | 51 / 56 / 61 / 66 / 71 | 44 / 47 / 50 / 53 / 56 | 70 / 73 / 76 / 79 / 82 |
| ASE2 | 181 / 186 / 191 / 196 / 201 | 138 / 141 / 144 / 147 / 150 | 164 / 167 / 170 / 173 / 176 |
| ASE3 | 321 / 326 / 331 / 336 / 341 | 250 / 253 / 256 / 259 / 262 | 201 / 204 / 207 / 210 / 213 |
| ASE4 | 442 / 447 / 452 / 457 / 462 | 346 / 349 / 352 / 355 / 358 | 315 / 318 / 321 / 324 / 327 |
| ASE5 | 524 / 529 / 534 / 539 / 544 | 416 / 419 / 422 / 425 / 428 | 350 / 353 / 356 / 359 / 362 |
| Orden por fila | RECIPROCIDAD, ENEL, ENERBIT, OCCIDENTE, CIUDAD LIMPIA-ACUEDUCTO | mismo | mismo |

**Valores golden Q1 de referencia (caché golden, bloque R1 `D9:D14` por hoja — para A2):**

| Empresa | ASE1 | ASE2 | ASE3 | ASE4 | ASE5 | TOTAL |
|---|---|---|---|---|---|---|
| RECIPROCIDAD EAAB | 0 | 13570060 | 0 | 7264980 | 0 | 20835040 |
| ENEL | 16180195504.29 | 11762458070.29 | 9648710791.68 | 10318956340.32 | 8229736739.09 | 56140057445.67 |
| ENERBIT | 0 | 0 | 0 | 0 | 126372775.04 | 126372775.04 |
| OCCIDENTE Directa | 524136930.28 | 381752310.90 | 453277723.14 | 226188183.51 | 163200815.15 | 1748555962.98 |
| EAAB-CL | 0 | 0 | 0 | 0 | 0 | 0 |
| **Σ = TOT_OPT (D9:D13)** | 16704332434.57 ✓ | 12157780441.19 ✓ | 10101988514.82 ✓ | 10552409503.83 ✓ | 8519310329.28 ✓ | — |

(Tablas equivalentes R2/R4 por empresa las extrae T0-0.5 del caché golden; este plan no las inventa.)

### 2.2 Decisiones de arquitectura

| ID | Opción elegida | Descartada | Por qué |
|---|---|---|---|
| D1 | Escribir **operandos** de las fórmulas detalle (filas `Total` por empresa zona principal + celdas leaf análogas en R2/R4), nunca las filas detalle con `<f>` | Escribir las filas detalle (r56, r66…) directamente | `F56=F33+F14-L14` es fórmula: escribirla rompería la cadena (violación §6 Rector). El guard `cell.CellFormula is not null` existente ya lo impide; T0 identifica los operandos reales |
| D2 | Misma pasada de escritura HU-07 (overload lista; un `File.Copy` + validación pre/post con mapa ampliado) | Segundo pase / segundo `File.Copy` / writer separado | Las `REMUNERACION_*` no se escriben (V3); la superficie nueva son más celdas leaf en la misma sesión. Un segundo pase rompería la atomicidad certificada y el hash A4 |
| D3 | `WorkbookLeafCellMapPorEmpresa`: dict explícito (empresa × `Ase.Id` → refs editables + visibles esperados), hermano de `WorkbookLeafCellMapPorAse` | Reutilizar offsets aritméticos entre empresas o bloques | Layouts heterogéneos (col C vs B vs A; EAAB-CL con menos fórmulas; ceros estáticos). Riesgo §11.3 Rector |
| D4 | Gate Σ empresas = visible por ASE/hoja en `ValidadorBasico` + `WorkbookLeafCoherence`, tolerancia ±0.5, ceros legítimos | Exigir no-cero por empresa o validar solo totales | V8: matriz dispersa real (EAAB-CL todo 0 en Q1). No-cero rompería Q1 |
| D5 | `Recaudo *`: modelo `RecaudoEmpresaInputs` (valores por concepto/ASE) + escritura en valores filas ~3–28; fila 29+ al mapa protegido | Tratar `Recaudo *` como derivadas | V9 + docx: son inputs pegados en valores con validaciones propias desde fila 29 |
| D6 | `VALIDACION_*`/`GERENTES_*`/`REMUNERACION_*`/`REPORTE…`/`BCE…`/`DetRetri…` al mapa de fórmulas protegidas (fallan la escritura si alguna deja de ser fórmula) | Ignorarlas en la validación pre/post | Blindan 2.3–2.7 contra escritura accidental desde 2.2 |
| D7 | Fuente primaria por hoja: R1-detalle ← `Recaudoporcomponente` (V10); `Recaudo *` ← `RecaudosReversados` (V11); R2/R4-detalle ← T0 decide entre fuente R2/R4 vs `RecaudosReversados`/`ReportePagosxBanco`/`Balance`/`Conciliaciones` | Fijar fuente de R2/R4/`Recaudo *` sin evidencia | V12 muestra que `Conciliaciones` no es copia directa; la dirección se demuestra, no se supone |
| D8 | Sin cambios funcionales UI: resumen por empresa en `txtLog` + Serilog (una línea por empresa×ASE con esperado post-Excel) | Grid por empresa / selectores / restyle | Delta mínimo; OPA sigue siendo Ejecutar |

### 2.3 Escritura 2.2 (misma sesión atómica)

```text
File.Copy plantilla → salida (una vez, igual que HU-07)
  └─► ValidarFormulasProtegidas (mapa HU-07 + mapa 2.2: detalle-empresa visibles,
      REMUNERACION_*×5, VALIDACION_*×5, GERENTES_*×5, Recaudo * fila 29+, resto 2.3–2.7)
        └─► EscribirCeldasLeaf HU-07 (bloques, intacto)
        └─► EscribirCeldasEmpresa 2.2 (operandos por empresa×ASE + Recaudo * valores)
              └─► revalidar fórmulas protegidas → guardar
```

Ante cualquier fallo: borrar salida parcial (patrón existente). Plantilla origen jamás mutada (hash A4 se mantiene y se extiende al mapa 2.2).

### 2.4 Dominio (Core, sin deps)

```csharp
public sealed class EmpresaFacturacion  // catálogo 2.2, nombres exactos de template
{
    // ReciprocidadEaab, Enel, Enerbit, OccidenteDirecta, EaabCiudadLimpia
    // + LabelTemplate ("ENEL", "OCCIDENTE", "RECIPROCIDAD", "ENERBIT", "CIUDAD LIMPIA-ACUEDUCTO")
}

public sealed class ConciliacionEmpresaInputs  // por ASE: 5 empresas × (R1 ops + R2 ops + R4 ops)
{
    public Ase Ase { get; set; } = new();
    public IReadOnlyDictionary<string, decimal> CeldasR1 { get; set; } = ...; // ref→valor, refs de T0
    public IReadOnlyDictionary<string, decimal> CeldasR2 { get; set; } = ...;
    public IReadOnlyDictionary<string, decimal> CeldasR4 { get; set; } = ...;
    // Visibles esperados por empresa (aritmética T0) para A2 + gate Σ
}

public sealed class RecaudoEmpresaInputs  // hojas Recaudo *: valores por concepto
{
    public string HojaRecaudo { get; set; } = ""; // "Recaudo ENEL", ...
    public IReadOnlyDictionary<string, decimal> Celdas { get; set; } = ...; // filas ~3-28, refs de T0
}
```

`WorkbookLeafInputs` suma `Conciliacion: IReadOnlyList<ConciliacionEmpresaInputs>` + `Recaudos: IReadOnlyList<RecaudoEmpresaInputs>` (listas vacías = comportamiento HU-07 intacto; compatibilidad hacia atrás por construcción). `IWorkbookLeafInputReader`: overload/método por empresa (contrato intacto, implementación extendida). `ProcesadorPeriodo`: pasos 2.2 integrados al flujo por ASE (leer→calcular→leaf+empresa→validar) con fail-fast que nombra ASE **y empresa**; una escritura al final (G2).

### 2.5 `IValidador` 2.2 (sin reabrir HU-04..07)

1. Todo lo HU-07 intacto (matcheo estricto por `Ase.Id`, `AjustesSfT=0` Q1, `GranTotal=Σ`, gates leaf-vs-consolidado).
2. Nuevo: por cada ASE y hoja (R1/R2/R4): `Σ visibles-empresa == visible de bloque` ±0.5 (R1: Σ = TOT_OPT `TotalOportunoEsperadoPorAse`; R2/R4 análogos con sus visibles T0). Ceros legítimos.
3. Nuevo: `Recaudo *` coherencia interna mínima según T0 (p. ej. TOTAL = suma ASE si el template lo formula; si son valores pegados, solo rango/positividad — T0 lo tipifica).
4. El validador NO abre `.xlsx` (igual que HU-06/HU-07).

### 2.6 UI — Visual Design Intent (delta mínimo)

Densidad Balanced, mismos GroupBoxes, sin restyle/colores/iconos. El `txtLog` agrega, por cada ASE, 5 líneas empresa (visible esperado post-Excel + fuente de la que se leyó) + Σ vs bloque. Serilog: mismos eventos HU-07 con propiedad `Empresa`. Sin nuevos controles (a lo sumo un `CheckBox` "Incluir conciliación por empresa", default ON según apruebe el Ingeniero; si OFF, el proceso es HU-07 puro — rollback funcional inmediato).

### 2.7 Golden Capa A extendida a 2.2 (honestidad HU-06/HU-07)

| # | Qué | Contra qué | Tol |
|---|---|---|---|
| A1 | Leaf 2.2 escritos en la **salida** (operandos empresa×ASE + `Recaudo *` valores) | Mismas celdas **leaf** del golden | ±0.5 |
| A2 | Visibles de **dominio** por empresa (aritmética T0 + Σ=bloque; tabla R1 §2.1; R2/R4 de T0-0.5) | Caché golden de detalle + `REMUNERACION_* D_*` | ±0.5 |
| A3 | `REMUNERACION_*`/`VALIDACION_*`/`GERENTES_*`/`Recaudo *` 29+ siguen siendo fórmula en la salida | Estructura | n/a |
| A4 | SHA256 plantilla origen igual antes/después | — | n/a |
| A5 | **Prohibido** comparar caché de fórmula de la salida vs golden; **prohibido** usar `TotOpt` HU-02 como visible por empresa | — | prohibido |
| A6 | `Conciliaciones` y `R*_Remuneracion` NO son oráculo de merge (control/manual) | — | n/a |

Capa B (manual Excel: abrir, recalcular, comparar `REMUNERACION_*` vs golden) fuera de CI, protocolo §5.3.

### 2.8 File changes previstos

| Archivo | Acción | Motivo |
|---|---|---|
| `Remuneracion.Core/Models/EmpresaFacturacion.cs` | Crear | Catálogo 5 empresas + labels de template |
| `Remuneracion.Core/Models/ConciliacionEmpresaInputs.cs` | Crear | Operandos + visibles por empresa×ASE |
| `Remuneracion.Core/Models/RecaudoEmpresaInputs.cs` | Crear | Valores hojas `Recaudo *` |
| `Remuneracion.Core/Models/WorkbookLeafInputs.cs` | Modificar | Sumar `Conciliacion` + `Recaudos` (listas, default vacías) |
| `Remuneracion.Core/Interfaces/IWorkbookLeafInputReader.cs` | Modificar | Método lectura por empresa (overload) |
| `Remuneracion.Core/Interfaces/IValidador.cs` | Modificar | Overload/gate Σ empresas (§2.5) |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | Modificar | Gate Σ por ASE/hoja, ceros legítimos |
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` | Modificar | Pasos 2.2 por ASE, fail-fast nombra ASE+empresa |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapPorEmpresa.cs` | Crear | Mapa explícito empresa×ASE (congelado T0) |
| `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs` | Modificar | Lectura por empresa (labels col C fuente R1, etc. según T0) |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs` | Modificar | Coherencia Σ empresas = visible por bloque |
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs` | Modificar | Escribir celdas 2.2 en la misma pasada + mapa protegido ampliado |
| `Remuneracion.Infrastructure/FileSystem/ArchivoFuenteLocator.cs` | Modificar **solo si** T0 demuestra input nuevo (`Conciliaciones/`) | Resolver ruta control/input; si no, intacto |
| `Remuneracion.WinForms/Form1.cs` (+Designer si aplica checkbox) | Modificar | Resumen por empresa en log; checkbox opt-out |
| `Remuneracion.IntegrationTests/GoldenConciliacionEmpresaTests.cs` | Crear | Capa A 2.2 (§2.7) |
| `Remuneracion.IntegrationTests/ConciliacionEmpresaTests.cs` | Crear | Dominio: gate Σ, ceros legítimos, mismatch nombra empresa |
| `Remuneracion.IntegrationTests/ProcesadorPeriodoTests.cs` | Modificar | Casos 2.2 (integración período con empresa) |

**No tocar (salvo bug blocker):** `IPlantillaWriter`/validation-only (HU-04); `IRecaudoReader` agregados HU-02; coherencia `F25`-Extemp HU-05; semántica single-ASE; `DetRetriRounder` (2.6); `requirements/` legado.

---

## 3. SPEC / Acceptance Criteria (trazabilidad Propuesta §10)

### Requirement 1 — Lectura por empresa (CA-1)

El sistema **MUST** leer el desglose por empresa de los 5 ASE desde sus reportes (R1: filas col C empresa + col D Total, V10). **MUST NOT** asumir labels de ASE1 para otros bloques sin verificación T0; si un label falta → fallo que nombra ASE+empresa, nunca valor inventado.

- GIVEN carpeta `1-Promoambiental` → WHEN lectura empresa R1 → THEN ENEL `F=19104244534.3`-equivalente y OCCIDENTE por componente ±0.5 vs fuente.
- GIVEN empresa sin recaudo (EAAB-CL Q1) → THEN ceros explícitos (no ausencia).

### Requirement 2 — Coherencia por empresa (CA-2)

El sistema **MUST** cumplir Σ empresas = visible de bloque por ASE y hoja (R1/R2/R4) ±0.5, con ceros legítimos. **MUST NOT** presentar agregados HU-02 como valores por empresa.

- GIVEN leafs 2.2 Q1 → THEN tabla §2.1 exacta ±0.5 + equivalentes R2/R4 de T0-0.5.

### Requirement 3 — Escritura 2.2 preservando fórmulas (CA-3/CA-4)

El sistema **MUST** escribir solo celdas del mapa por empresa + `Recaudo *` valores en la misma pasada HU-07; **MUST NOT** escribir ninguna celda con `<f>` (incluye filas detalle r51–71…, `REMUNERACION_*`, `VALIDACION_*`, `GERENTES_*`, `Recaudo *` 29+); **MUST NOT** mutar la plantilla; ante fallo **MUST** borrar la salida parcial.

- GIVEN 5 leafs + empresa válidos → WHEN overload lista → THEN cambian solo celdas del mapa; A3 verde.
- GIVEN gate Σ roto en ASE3–OCCIDENTE → THEN `CalculoInvalidoException` que nombra ASE+empresa, sin archivo certificado.

### Requirement 4 — Hojas `Recaudo *` en valores (CA-3)

El sistema **MUST** diligenciar las 5 hojas `Recaudo *` (zona datos ~3–28) con valores de fuente según mapa T0; fila 29+ intacta como fórmulas.

### Requirement 5 — Trazabilidad y UI honesta (CA-6/CA-7)

Serilog + `txtLog` por ASE **y por empresa** (qué fuente→qué celdas, Σ vs bloque). Resumen etiqueta "esperado post-Excel". Sin restyle.

### Requirement 6 — Golden Capa A 2.2 (CA-3/CA-5)

Matriz §2.7 para 5 empresas × 5 ASE con insumos Q1. **MUST NOT** comparar caché de salida vs golden (A5).

| CA §10 | HU-08 |
|---|---|
| CA-1 | Lee desglose por empresa de los 5 ASE (fail-fast nombra ASE+empresa) |
| CA-2 | Σ empresas = visible por bloque (R1/R2/R4); agregados HU-02 ≠ valores empresa |
| CA-3 | Capa A 2.2; `REMUNERACION_*` correctas post-Excel (Capa B manual residual) |
| CA-4 | Reassert mapa ampliado + plantilla no mutada (hash) |
| CA-5 | Gate Σ + coherencia leaf por empresa |
| CA-6 | Serilog + log por empresa |
| CA-7 | Mismo flujo 5-ASE + resumen por empresa |

---

## 4. TASKS

## Review Workload Forecast

| Field | Value |
|---|---|
| Estimated changed lines | 800–1300 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 T0 + mapa → PR2 lectura por empresa → PR3 writer + mapa protegido → PR4 validador + coherencia Σ → PR5 período + UI + golden |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes. Chained PRs recommended: Yes. Chain strategy: pending.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|---|---|---|---|
| 0 | Discovery T0 + mapa congelado | PR 1 | Bloquea todo; solo lectura + datos |
| 1 | Lectura por empresa (reader) | PR 2 | Depende de PR 1 |
| 2 | Writer 2.2 + protegido ampliado | PR 3 | Depende de PR 1 |
| 3 | Validador Σ + coherencia | PR 4 | Depende de PR 1 |
| 4 | Período + UI + golden 2.2 | PR 5 | Depende de PR 2–4; 25 tests previos verdes |

### Phase 0 — Discovery T0 (BLOQUEANTE, solo lectura)

- [ ] 0.1 R1: por cada bloque ASE, dumpear operandos de cada fórmula detalle-empresa (`F{56,66,…}=…`) y clasificar valor-vs-fórmula; registrar filas `Total` por empresa de la zona principal (análogas a r14/r33) con sus labels exactos de template.
- [ ] 0.2 R2: idem por fila empresa col B (44/47/…): fórmula visible del bloque (análoga a `E41=E15+E26-K15`) + operandos por empresa + clasificación.
- [ ] 0.3 R4: idem col A (70/73/…): fórmula por empresa + operandos + clasificación.
- [ ] 0.4 Fuentes R2/R4 por ASE: ¿traen desglose por empresa? (labels y columnas). Si NO: probar `RecaudosReversados` / `ReportePagosxBanco` / `Balance` como origen; documentar el mapeo o declarar el hueco.
- [ ] 0.5 Extraer del caché golden las tablas por empresa R2/R4 (equivalente a §2.1) + explicar EAAB-CL 145/70/70 vs 150/75/75 + verificar `GERENTES_*` = fórmulas puras + barrido de externalLinks en las 5 hojas.
- [ ] 0.6 `Recaudo *`: dumpear zona datos (~3–28: ¿una fila por ASE? ¿columnas por concepto?) + fila 29+ (fórmulas a proteger) en las 5 hojas; mapear cada celda a `RecaudosReversados` (bloque 66/ASEO + detalle banco) o a `Conciliaciones` (fijar dirección con la discrepancia V12 como test).
- [ ] 0.7 Congelar `WorkbookLeafCellMapPorEmpresa` (empresa × ASE → editables + visibles + hojas `Recaudo *`). **Nada entra al código sin esta tabla.**

### Phase 1 — Dominio (modelos + lectura)

- [ ] 1.1 `EmpresaFacturacion` (labels exactos: `ENEL`, `OCCIDENTE`, `RECIPROCIDAD`, `ENERBIT`, `CIUDAD LIMPIA-ACUEDUCTO`) + `ConciliacionEmpresaInputs` + `RecaudoEmpresaInputs`; extender `WorkbookLeafInputs` (listas default vacías).
- [ ] 1.2 Reader por empresa (R1/R2/R4 + `Recaudo *`) según T0-0.7; labels dinámicos, `ServEspK=0` donde aplique; fallo nombra ASE+empresa.
- [ ] 1.3 `ArchivoFuenteLocator`: solo si T0-0.4/0.6 exige nuevo input (si no, intacto).

### Phase 2 — Escritura (misma pasada)

- [ ] 2.1 `WorkbookLeafCellMapPorEmpresa.cs` (datos T0-0.7) + escritura de celdas 2.2 dentro del overload lista existente.
- [ ] 2.2 Mapa protegido ampliado (`REMUNERACION_*`, `VALIDACION_*`, `GERENTES_*`, `Recaudo *` 29+, resto 2.3–2.7; shared-formula awareness HU-07T0 §0.5 para D106 y `G56:AP56`-análogos).
- [ ] 2.3 Borrado de parcial + hash plantilla intactos.

### Phase 3 — Orquestación + UI delta mínimo

- [ ] 3.1 `ProcesadorPeriodo`: pasos 2.2 por ASE (leer empresa → leaf → validar Σ), fail-fast ASE+empresa, una escritura.
- [ ] 3.2 `Form1`: resumen por empresa en log (+ checkbox opt-out opcional, default ON).
- [ ] 3.3 Serilog por empresa (propiedad `Empresa`, mismos sinks).

### Phase 4 — Pruebas y evidencia

- [ ] 4.1 `ConciliacionEmpresaTests` (in-memory): Σ ok 5/5, mismatch nombra empresa (ej. ASE3–OCCIDENTE), ceros EAAB-CL legítimos, no-cero indebido falla, GranTotal intacto.
- [ ] 4.2 `GoldenConciliacionEmpresaTests`: matriz §2.7 (5 empresas × 5 ASE, fixtures Q1; `Conciliaciones/` y `R*_Remuneracion` NO son oráculo).
- [ ] 4.3 Integración período con empresa + negativa (falta label empresa en fuente → ASE+empresa en el error, sin salida).
- [ ] 4.4 Los 25 tests existentes verdes; build 0 warnings; CRLF; sin commit.

### Phase 5 — Documental

- [ ] 5.1 Capa B manual §5.3 ejecutada una vez y evidenciada (sin fingirla como gate de merge).
- [ ] 5.2 Cierre deja explícito el frente 2.3 (siguiente HU propuesta: `REPORTE RECAUDO x BANCO`).

---

## 5. Test Strategy

| Capa | Qué | Cómo |
|---|---|---|
| Unidad | Gate Σ empresas = bloque (R1/R2/R4 × 5 ASE, ceros legítimos) | In-memory, tabla §2.1 + T0-0.5 |
| Unidad | Reader por empresa (labels dinámicos, ASE4-sin-Especiales análogo) | Fuentes Q1 reales, sin Excel de salida |
| Integración | Período con 2.2 (5 carpetas reales, salida temp) | Insumos Q1, fail-fast ASE+empresa |
| Golden Capa A 2.2 | Matriz §2.7 | OpenXML read-only + aritmética dominio; plantilla=golden (A5 aplica) |
| UI | Resumen por empresa | Funcional manual (sin harness) |
| Capa B | `REMUNERACION_*` post-Excel | Manual — §5.3 |

### 5.1 Fixtures

- Golden/plantilla: `Docs/Insumos/Remuneracion 202607-1 Total.xlsx`.
- Fuentes: `Docs/Insumos/REMUNERACION 2026071/{1..5}-*/` (R1/R2/R4 + `RecaudosReversados` + `ReportePagosxBanco` + `Balance`); `Consolidado/Conciliaciones/` y `Consolidado/Remuneración del sistema/` como **control**, no oráculo (V12/V13).
- Referencia: tabla §2.1 (R1) + T0-0.5 (R2/R4), tolerancia ±0.5.

### 5.2 Casos negativos obligatorios (nombran empresa y ASE)

Falta label `ENERBIT` en fuente R1 ASE5; Σ OCCIDENTE ≠ bloque en ASE3; `Recaudo ENEL` con zona datos incompleta; salida == plantilla (no in-place); overwrite cancelado; EAAB-CL con valor inesperado ≠ 0 (alerta, no necesariamente fallo — T0 tipifica); Q2 en `Conciliaciones` (= 0, sin efecto en Q1).

### 5.3 Protocolo manual Capa B (no CI)

1. Generar salida a ruta distinta del golden. 2. Abrir en Excel, recalcular. 3. Comparar `REMUNERACION_*` bloques R1/R2/R3/R4/consolidado vs golden ±0.5 (refs §2.1 + T0-0.5) y `Recaudo *` vs golden. 4. Evidencia adjunta. No es gate de merge.

---

## 6. Architecture Validation Certificate

### 6.1 SOLID

| Principio | Veredicto |
|---|---|
| **S** | Empresa = extensión de datos (mapa + inputs + reader + gate); período/orquestación/writer conservan su rol HU-07. |
| **O** | Se agregan modelos/mapas/overloads; el path HU-07 funciona con listas 2.2 vacías (abierto sin modificar). |
| **L** | `OpenXmlPlantillaWriter` escribe más celdas del mismo modo leaf; comportamiento HU-07 puro inalterado. |
| **I** | Catálogo `EmpresaFacturacion` separado de `Ase`; reader/validador crecen por overload. |
| **D** | Core define empresa/inputs/gates; Infrastructure/WinForms componen. Sin nuevas deps. |

### 6.2 Best Practices

- La verdad del workbook manda: `REMUNERACION_*` 100 % fórmulas ⇒ cero escritura directa (Rector §6 preservación).
- Mapa explícito empresa×ASE verificado, no offsets (Rector §11.3).
- Una escritura atómica; plantilla nunca mutada; hash A4 extendido.
- Golden honesto por empresa (OpenXML no recalcula; A5/A6).
- Fail-fast nombra ASE+empresa; sin salida certificada ante fallo.
- Docx leídos como spec funcional (no como folklore): citas V9/V11 + Inst. 9 (`Informe AFaseo` se ignora).

### 6.3 Performance

- +25 filas detalle R1 + R2/R4 análogas + 5×~26 celdas `Recaudo *` ≈ <300 escrituras extra en la misma sesión OpenXML. Irrelevante a esta escala; `Task.Run` existente para no congelar el form.

**Veredicto:** APROBADO como it. 2.2 de Fase 2 **si** T0 congela el mapa con evidencia y se acepta CA-3 parcial (Capa A en CI, Capa B manual).

---

## 7. Riesgos

| Riesgo | Likelihood | Mitigación |
|---|---|---|
| Fuente R2/R4 no trae desglose por empresa y el origen real es otro archivo | Alta | T0-0.4 lo decide antes de codificar; si el hueco es real, la HU se recorta a R1+`Recaudo *` con rebase (no se inventa fuente) |
| `Conciliaciones/` resulta ser input obligatorio (dirección V12 invertida) | Media | T0-0.6; si entra, `ArchivoFuenteLocator` se extiende y el plan se rebasea (Unidad 0) |
| Algún operando empresa es fórmula no editable (caso `F56` a la inversa) | Media | T0-0.1–0.3 clasifican antes de prescribir; el guard `<f>` existente protege |
| EAAB-CL 145/70/70 es layout distinto (no solo ceros) | Media | T0-0.5; el mapa por empresa absorbe la diferencia sin offsets |
| Ceros legítimos confundidos con "falta de lectura" | Alta | Gate Σ con ceros explícitos + test EAAB-CL Q1; el reader distingue "leído 0" de "label ausente" (fallo) |
| Inflar a 2.3–2.7 dentro de esta HU (`VALIDACION_*` tienta) | Media | §0.2 out-of-scope + D6 (protegidas, no implementadas); rechazar PRs que lo metan |
| Comparar caché de salida vs golden y "cerrar" CA-3 en falso | Alta | A5 en tests; este plan lo prohíbe |
| Carpeta `5-Área Limpia` (tilde) en nuevos tests | Media | Reutilizar `Insumos.cs` (ya lo maneja) |

---

## 8. Rollback

- Revertir PRs en orden inverso (golden → período/UI → validador → writer → reader → mapa/T0).
- HU-04..HU-07 intactas sin esta HU: listas 2.2 vacías = comportamiento HU-07 puro (más checkbox opt-out si se implementa).
- `Docs/Insumos/` untracked: jamás como destino; pisada se restaura desde backup.
- Si T0 demuestra hueco de fuente (Riesgo 1/2), recorte con rebase, no invención.

---

## 9. Decisiones fijadas por este plan

1. Rector = Propuesta §9 it. 2.2 / §6 / §10 (±0.5). 2.3–2.7 y Q2 quedan fuera, citados a §9.
2. `REMUNERACION_*` = 100 % fórmulas: no se escriben; se protegen. Superficie 2.2 = operandos empresa en R1/R2/R4 + `Recaudo *` en valores.
3. Una sola escritura atómica (overload lista HU-07); sin segundo pase.
4. Gate Σ empresas = visible por bloque y ASE (R1/R2/R4, ±0.5, ceros legítimos) — coherencia central 2.2, verificada 5/5 en R1 contra golden.
5. Mapa explícito empresa×ASE congelado por T0; ninguna dirección sin evidencia.
6. Golden Capa A 2.2 con honestidad HU-06/HU-07; Capa B manual residual; `Conciliaciones`/`R*_Remuneracion` = control, no oráculo.
7. UI delta mínimo + Serilog por empresa; OPA = Ejecutar.
8. Apply espera aprobación + PRs encadenados (Unidad 0–4, §4).

**Listo para aprobación del Ingeniero. No implementar hasta OK.**
