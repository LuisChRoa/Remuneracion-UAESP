# T0 — Geometría congelada de la plantilla estandarizada única (Plan 38)

> **Fixture que manda:** `Remuneracion.IntegrationTests/GeometriaPlantillaEstandarizadaT0Tests.cs`
> (reutiliza `ContratoSlotR1T0Tests.cs`, no lo duplica). Lectura SOLO con BCL
> (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`); sin Excel/COM, sin
> DocumentFormat.OpenXml, sin ExcelDataReader.
> **Oráculo geométrico:** `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`
> (período interno base **2026092**). Los números citados aquí son citas del disco, no autoridad:
> el test compara código-vs-fixture y falla ante cualquier deriva.

## 1. Frontera D-B: firma (derivado en runtime) vs estructural (congelado)

| Clase | Qué es | Tratamiento |
|---|---|---|
| **Firma (derivado)** | Celdas-dato del interior del bloque `Reporte Componentes R1` (filas 6-72/110-210/289-355/450-550/603-676) | Se derivan en runtime por firma A–E + clasificador de frontera + sinonimia (`EscritorValoresSlotR1`). NO se congelan por celda. |
| **Estructural (congelado)** | Cadenas `CONSOLIDADO` D, visibles de bloque R1/R2/R4, refs `DetRetri`/`DetValiRetri`, visibles `SALDOS POR NOTA`/`RETRIBUCION NEGATIVA`/`AJUSTES - SF-T` | Congelado como DATOS tabulares (evidencia del disco) con drift gate. Sin firma resoluble → congelar es correcto (D-E). |

## 2. Cadenas congeladas (cita celda-a-celda, dump 2026-10-14)

### 2.1 CONSOLIDADO_TOTAL RECAUDO (columna D)
| Celda | Fórmula |
|---|---|
| D9:D13 | `'Reporte Componentes R1'!F75 / F254 / F415 / F568 / F680` |
| D28:D32 | `'Rem. Anticipos R2'!E47 / E144 / E262 / E362 / E446` |
| D47:D51 | `'Reporte Componentes R1'!F77 / F256 / F417 / F570 / F682` |
| D66:D70 | `'Reversion Pagos R4'!D73 / D174 / D211 / D326 / D361` |
| D85:D89 | `'AJUSTES - SF-T'!D47 / D48 / D49 / D50 / D51` |
| D104:D108 | `D9+D28+D47+D66+D85` … `D13+D32+D51+D70+D89` (D106 = shared follower del maestro D104) |
| D109 | `SUM(D104:D108)` |
| D111 | `D109+N109` |

### 2.2 Visibles `Reporte Componentes R1` por ASE
| ASE | TOT_OPT (→ D9:D13) | EXTEMP (→ D47:D51) |
|---|---|---|
| 1 | `F75 = F47+F70+F16-L16-L47` | `F77 = F59+F32-L32` |
| 2 | `F254 = F171+F208-L171+F120-L120` | `F256 = F190+F146-L146` |
| 3 | `F415 = F330+F353+F299-L299-L330` | `F417 = F342+F315-L315` |
| 4 | `F568 = F511+F548+F460-L460-L511` | `F570 = F530+F486-L486` |
| 5 | `F680 = F674+F648+F613-L613-L648` | **`F682` = VALOR 0 (no fórmula)** |

### 2.3 Visibles `Rem. Anticipos R2` / `Reversion Pagos R4` por ASE
| ASE | R2 visible | R4 visible |
|---|---|---|
| 1 | `E47 = E21+E32-K21` | `D73 = D15-P15` |
| 2 | `E144 = E94+E112-K94` | `D174 = D111-P111` |
| 3 | `E262 = E182+E193-K182` | `D211 = D206-P206` |
| 4 | `E362 = E309+E327-K309` | `D326 = D249-P249` |
| 5 | `E446 = E418+E404-K404` | `D361 = D358-P358` |

### 2.4 DetRetri2026092 / DetValiRetri2026092
- DetRetri D23:D28 = diferencias vs `CONSOLIDADO D104:D109`; D32:D36 = booleanos de composición
  (`+D23+'Reporte Componentes R1'!F75+'Rem. Anticipos R2'!E47+'Reversion Pagos R4'!D73+'Reporte
  Componentes R1'!F77+'AJUSTES - SF-T'!D47=D9`, …).
- DetValiRetri D16:D21 = diferencias vs `CONSOLIDADO U104:U109`; D24:D29 = booleanos
  (`+'Reporte Componentes R1'!M75+'Reporte Componentes R1'!M77+'Rem. Anticipos R2'!L47+'Reversion
  Pagos R4'!J73+D16+'AJUSTES - SF-T'!U47=D9`, …).

### 2.5 SALDOS POR NOTA / RETRIBUCION NEGATIVA / AJUSTES - SF-T
- SALDOS visibles `C10/C23/C36/C49/C62 = C8-I8 / C21-I21 / C34-I34 / C47-I47 / C60-I60`.
- RETRIBUCION visibles idénticas (`C10/C23/C36/C49/C62`).
- AJUSTES `D9:D13 = 'SALDOS POR NOTA'!C10/C23/C36/C49/C62`; `D28:D32 = 'RETRIBUCION NEGATIVA'!C10/C23/C36/C49/C62`;
  `D47 = D9+D28`; `D52 = SUM(D47:D51)`.

### 2.6 Familias SIN cambio de geometría (identidad verificada byte-a-byte)
- **BCE SC POR FACT.**: `F3:F7 = D+E`, `H3:H7 = DetRetri2026092!J9..J13`, `I3:I7 = F-H`,
  `F9 = SUM(F3:F7)`, `H9 = SUM(H3:H7)`, `H10 = H9=DetRetri2026092!J14`, `D11/E11/F11`, `F12/F13`,
  `E19:E24/F19:F24/G19:G24/H19:H24`.
- **INTERVENTORIA**: `K31/M31/N31 = SUM(K26:K30 / M26:M30 / N26:N30)`, `K32 = SUM(M31:N31)`.
- **ANT EXT-REV**: `C3 = C11+C19+C27+C35+C43`, `M3/N3/C5`.
- **VALIDACION_* / VALIDACION_TOTAL / GERENTES_***: `O/P` y `D14 = SUM(D9:D13)`.
- **REMUNERACION_***: bloques `D9:D13/D14/D28:D32/D33/D47:D51/D52/D66:D70/D71/D90/D104:D108/D109`
  (fragmentos por NOMBRE de hoja, robustos a la fila). Única excepción: las filas FÓRMULA del bloque
  AJUSTES `D85:D89` por empresa (Reciprocidad 86/88, Enel 85/87/89, **Enerbit 89**, Occidente y
  EaabCL ninguno) — ver `WorkbookLeafCellMapPorEmpresa.AjustesFormulaPorEmpresa`.
- **REPORTE RECAUDO x BANCO**: la hoja termina en la fila 79 → **NO existe** `C81`/`D81` (el legacy
  los declaraba; se retiraron del mapa protegido).

## 3. Drift gate (R-DRIFT-FUTURO)

El fixture `GeometriaPlantillaEstandarizadaT0Tests` falla si la plantilla cambia en:
1. Inventario de hojas (nombres + orden + conteo 40) — `T0_DriftGate_WorkbookXmlYConteos`.
2. Base interna `N3 = 2026092`.
3. Cualquier cadena celda-a-celda de §2.

**Procedimiento ante un cambio de plantilla (por el negocio):** no se edita ningún literal a mano.
Se ejecuta un **nuevo T0** (re-lectura BCL de los bytes), se actualiza este doc + el fixture, y el
diff queda auditado por el Ingeniero. El gate convierte el cambio silencioso en FAIL accionable.

## 4. Residuos declarados de este T0/T1/T2

- **Per-empresa Q1** (`EditablesR1PorEmpresa` / `EditablesR2PorEmpresa` / `EditablesR4PorEmpresa`):
  los mapas de la conciliación por empresa (HU-08) siguen anclados a geometría legacy. Requieren su
  propia derivación por (empresa × ASE) desde el disco (la presencia de empresas por ASE difiere del
  legacy: p. ej. R2/R4 ASE2 traen `NUEVO ESQUEMA` y no `RECIPROCIDAD`). **Owner:** T3/T5 con su
  propio mini-T0 (R-ATRIBUCION: no se migra a ciegas).
- **Bloques editables SALDOS/RETRIB** (`FilasSaldosNotasPorAse` / `FilasRetribucionNegativaPorAse` /
  `Conceptos*PorAse`) y **DetalleR2R4** (`BloquesR2`/`BloquesR4`): pendientes de re-congelar.
- **R1 interior (HU-07 legacy)**: `R1Q2EditablesPorAse` y las entradas R1 de
  `EditableLeafCellsPorAse` quedan como el path legacy HU-07, SUPERSEDED por `EscritorValoresSlotR1`
  (la hoja R1 se gobierna por slots; esas entradas son código muerto en el path espejo).

## 5. T2-continuación — mini-T0 por (empresa × ASE) y bloques detalle (dump BCL 2026-10-14)

> Fixture: `GeometriaPlantillaEstandarizadaT0Tests.T0_PorEmpresaR2R4_DireccionesMatcheanFormulaOportuno`,
> `T0_DetalleR2R4_BloquesMatcheanPlantilla`, `T0_R1Q2Editables_DireccionesExistenYSonValores`.
> Cita de valores vs el manual legacy (rol↔celda): `Remuneracion.IntegrationTests/GeometriaGoldenLegacy.cs`.

### 5.1 Inventario de empresas por ASE (manda la FUENTE Q1, no el legacy)

| ASE | R2 (fuente) | R4 (fuente) |
|---|---|---|
| 1 | ENEL, OCCIDENTE | OCCIDENTE |
| 2 | ENEL, NUEVO ESQUEMA, OCCIDENTE | ENEL, NUEVO ESQUEMA |
| 3 | EAB, ENEL, OCCIDENTE | ENEL, OCCIDENTE |
| 4 | ENEL, NUEVO ESQUEMA, OCCIDENTE | ENEL, NUEVO ESQUEMA, OCCIDENTE |
| 5 | ENEL, ENERBIT, OCCIDENTE | ENEL, OCCIDENTE |

**Hallazgo R-ATRIBUCION:** el inventario de claves `(empresaId, aseId)` del legacy era correcto
(match con la fuente); lo que cambió fue la DIRECCIÓN. La única ausencia: el template R2/R4 ASE2
trae la fila `NUEVO ESQUEMA`/`RECIPROCIDAD` (104/83) pero la fuente Q1 de ASE2 no la provee → la
celda queda en su 0 de template (no se migra una empresa ausente; se reporta).

### 5.2 Direcciones nuevas (celdas-dato; operando de la fórmula OPORTUNO del template)

- **R4** (`EditablesR4PorEmpresa`, Total/P): ASE1 OCC D14/P14; ASE2 NUEVOESQ D104/P104, ENEL D98/P98;
  ASE3 ENEL D199/P199, OCC D205/P205; ASE4 NUEVOESQ D242/P242, ENEL D236/P236, OCC D248/P248;
  ASE5 ENEL D351/P351, OCC D357/P357. (La fórmula del bloque `D<n>=D<n>-P<n>` referencia la misma D.)
- **R2** (`EditablesR2PorEmpresa`, Total/SubsCont/Especiales): derivadas de la fórmula OPORTUNO por
  empresa (p. ej. ASE1 ENEL `E53=E24+E10+E6-K6-K10` → Total E10, SubsCont E24, Esp K10; ASE5 la
  fórmula legacy invierte el orden pero el ROL se fija por rótulo, no por orden).
- **R1-Q2** (`R1Q2EditablesPorAse`): operandos de `TOT_OPT`/`EXTEMP` del template (p. ej. ASE1
  `F75=F47+F70+F16-L16-L47` → Mes0/1/2 = F16/F47/F70; `F77=F59+F32-L32` → Aplic0/1 = F32/F59). El
  rol cuyo `L` NO aparece en la fórmula es el "último" Mes (Mes2), coherente con `Σ Mes − Esp(≤última)`.
  ASE5 trae 2 filas Mes (F613/F648) y su `EXTEMP` (F682) es VALOR 0.
- **Bloques detalle** (`BloquesR2`/`BloquesR4`): R2 = 3-33 / 67-113 / 164-194 / 282-328 / 382-419;
  R4 = 3-15 / 93-111 / 194-206 / 231-249 / 346-358 (primera fila = `Vlr Servicio`, última = `Total`).

### 5.3 Verificación numérica (R-ATRIBUCION, no asumida)

Para cada ASE se comprobó que `Σ VisibleR2(empresa) = visible de bloque` (p. ej. ASE2:
28506780.26 + 29406119.09 + 11318987.49 = 69231886.84 = `E144`) y que los operandos del leaf
coinciden ±0.5 con el manual legacy emparejado por ROL (`GeometriaGoldenLegacy`). El gate de
coherencia `ValidarSigmaEmpresas` (rol-based) queda intacto.

### 5.4 Residuos declarados de T2-continuación (NO migrados, con motivo)

- **`EditablesR1PorEmpresa` (R1 per-empresa)**: NO migrado. La fórmula OPORTUNO per-empresa del R1
  (`F85=F62+F36+F8-L8-L36`, …) referencia celdas que NO mapean 1:1 sobre el modelo de roles legacy
  (`FuenteR1.TotalMain/Medio/Menor/SubsMain/SubsMenor`) y el gate `ValidarSigmaEmpresas` es
  rol-based (no valida direcciones) → migrar sin el template legacy como evidencia arriesgaría
  R-ATRIBUCION. Sus direcciones legacy caen en celdas VALOR (sin `WRITE-EN-FORMULA`). **Owner:** T5
  con su propio mini-T0.
- **`EditableLeafCellsPorAse` (entradas R1) y `MapearR1` (celdas F25/F41/L25/F30/F10/L10)**: mismo
  caso (path legacy HU-07; celdas valor). Superseded por `EscritorValoresSlotR1`.
- **SALDOS/RETRIB** (`FilasSaldosNotasPorAse`/`ConceptosSaldosPorAse`): el template movió las filas
  (ASE1 3-8 con `Vlr Intereses` en 4, etc.); pendiente su re-congelamiento + los tests de
  `AjustesSfTTests` en la MISMA pasada (evitar nuevos rojos).
- **R1 gates** (`ValidadorTotalesR1Workbook` + fixtures `InterioresR1Esperados`/`TotalesR1Esperados`)
  y **shared masters** (`SharedFormulaR1Esperados`): freezan geometría legacy del manual;
  re-congelar o retirar con cita (T3/T5).

