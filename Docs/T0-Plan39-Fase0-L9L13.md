# T0 (Plan 39) — Phase 0: variantes L9:L13 por quincena + allow-list `<f>`

> **Fixture que manda:** `Remuneracion.IntegrationTests/GeometriaPlantillaEstandarizadaT0Tests.cs`
> (extiende el T0 del Plan 38; reutiliza `ContratoSlotR1T0Tests`, no lo duplica). Lectura SOLO con
> BCL (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`): sin Excel/COM, sin
> DocumentFormat.OpenXml, sin ExcelDataReader.
> **Oráculo geométrico:** `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`
> (período interno base **2026092**). Los números citados aquí son citas del disco (dump BCL
> 2026-10-14), no autoridad: el test compara código-vs-fixture y falla ante cualquier deriva.

## 1. R-INTERV-Q — `CONSOLIDADO_TOTAL RECAUDO!L9:L13` es quincena-dependiente por DOMINIO

Regla de primera clase (Plan 39 §2.2): si `Periodo.NumeroQuincena == 1` →
`L9:L13 = INTERVENTORIA!N26:N30` (Primera quincena); si `== 2` → `INTERVENTORIA!M26:M30`
(Segunda quincena). La quincena viaja como `Periodo`, **nunca** por contenido del workbook.

Texto-`<f>` EXACTO congelado (por celda; el oficial trae la forma Q2):

| Celda | `<f>` variante Q2 (oficial) | `<f>` variante Q1 | caché `<v>` Q2 | caché `<v>` Q1 |
|---|---|---|---|---|
| L9  | `INTERVENTORIA!M26` | `INTERVENTORIA!N26` | 189185988 | 189185987 |
| L10 | `INTERVENTORIA!M27` | `INTERVENTORIA!N27` | 266580256 | 266580255 |
| L11 | `INTERVENTORIA!M28` | `INTERVENTORIA!N28` | 154788536 | 154788535 |
| L12 | `INTERVENTORIA!M29` | `INTERVENTORIA!N29` | 120391083 | 120391083 |
| L13 | `INTERVENTORIA!M30` | `INTERVENTORIA!N30` | 128990446 | 128990445 |

**Notas normativas:**
1. ASE4 (L12) es IDÉNTICO en ambas variantes (120391083): el gate NO puede distinguir Q1/Q2 por
   ASE4 solo — la aserción es por **texto-`<f>`**, no por valor.
2. El off-by-one ASE1/2/3/5 (…988 vs …987) es el split de redondeo mensual `M+N = K ± 1`. Por eso la
   atribución es MATERIAL: pegar la columna contraria mueve ~1 por ASE.
3. Los cachés `<v>` solo **documentan** la variante; el oráculo es el texto-`<f>` (prohibido `<v>`
   como oráculo de fórmulas).
4. Oráculo de cachés por variante: `Remuneracion.Core.Constants.InterventoriaDeclarada`
   (`SegundaQuincenaPorAse` = M26:M30, `PrimeraQuincenaPorAse` = N26:N30) — se REUTILIZA, no se duplica.

## 2. Inventario de hojas de período (oficial)

| Posición | Hoja |
|---|---|
| 16 | `Informe AFaseo Recaudo 202609-2` |
| 35 | `DetRetri2026092` |
| 36 | `DetValiRetri2026092` |

`N3 = 2026092` (valor plano; último dígito = quincena). Phase 0 sella `N3 = periodo.CodigoCompleto`.

## 3. Las 11 `<f>` del token de período (superficie de reescritura de tokens)

| Hoja | Celda | `<f>` |
|---|---|---|
| `BCE SC POR FACT.` | H3  | `DetRetri2026092!J9` |
| `BCE SC POR FACT.` | H4  | `DetRetri2026092!J10` |
| `BCE SC POR FACT.` | H5  | `DetRetri2026092!J11` |
| `BCE SC POR FACT.` | H6  | `DetRetri2026092!J12` |
| `BCE SC POR FACT.` | H7  | `DetRetri2026092!J13` |
| `BCE SC POR FACT.` | H10 | `H9=DetRetri2026092!J14` |
| `DetValiRetri2026092` | K24 | `+DetRetri2026092!N32` |
| `DetValiRetri2026092` | K25 | `+DetRetri2026092!N33` |
| `DetValiRetri2026092` | K26 | `+DetRetri2026092!N34` |
| `DetValiRetri2026092` | K27 | `+DetRetri2026092!N35` |
| `DetValiRetri2026092` | K28 | `+DetRetri2026092!N36` |

Conteo exacto = **11** (0 extra). El parser de celdas del fixture respeta el autocierre
`<c .../>` (un regex ingenuo atribuye mal la fórmula a la celda vacía anterior).

## 4. Allow-list de `<f>` tocables SOLO por Phase 0 (R-T-2)

**16 celdas** = 11 tokens de período (§3) + 5 L9:L13 (§1). Cualquier otra `<f>` tocada por cualquier
fase = FAIL; el grep de cierre distingue fases por namespace/clase (`PreparadorBasePipeline` vs el
procesamiento values-only).

La reescritura de tokens de período también toca partes **NO-`<f>`** (nombres de hoja en
`xl/workbook.xml`, `docProps/app.xml`, `xl/sharedStrings.xml`, `definedNames`): es la copia mecánica
de período, fuera del alcance de esta allow-list de celdas-fórmula.

## 5. Drift gate

El fixture falla si el oficial cambia en: inventario de hojas (40, nombres + orden), base interna
`N3`, cualquier cadena celda-a-celda del T0 del Plan 38, o cualquiera de las dos variantes L de §1 /
las 11 `<f>` de §3 / el conteo de la allow-list (§4). Un cambio del oficial por el negocio exige un
**nuevo T0** (re-lectura BCL) + re-congelar AMBAS variantes L; nunca edición silenciosa de literales.

---

## 6. T3 — Sweep BCL a veredictos (Plan 39 §2.3)

> **Método:** barrido BCL (`System.IO.Compression.ZipArchive` + XML/regex) sobre la plantilla oficial
> `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` (base interna `2026092`).
> Sin Excel/COM/OpenXML. Lectura read-only. **Fecha:** 2026-10-14.
> **Hash del oficial (SHA256):** `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`
> (692.598 bytes; re-verificado intacto tras el sweep — R-INPLACE).
> **Superficie:** 40 hojas, 8.766 `<f>`. Cada hallazgo con veredicto citado; **cero supuestos
> silenciosos**. Evidencia de contraste decisiva: las 3 bases preparadas offline (2026071/2026072/
> 2026082, retiradas en T4) traen EXACTAMENTE las mismas referencias `INTERVENTORIA!M26:M31` que el
> oficial — el preparador offline NUNCA las conmutaba.

### 6.1 Veredictos por superficie

| # | Superficie | Hallazgo (celda citada) | Veredicto | Regla / motivo |
|---|---|---|---|---|
| S-MN-1 | Refs col M/N | `CONSOLIDADO_TOTAL RECAUDO!L9:L13 = INTERVENTORIA!M26:M30` | **Intencional-por-quincena — Phase 0 SÍ** | R-INTERV-Q: Q1→`N26:N30`, Q2→`M26:M30`. Única celda de selección de quincena; allow-list T0 (§4). |
| S-MN-2 | Refs col M/N | `INTERVENTORIA!M31=SUM(M26:M30)`, `N31=SUM(N26:N30)`, `K31=SUM(K26:K30)`, `K32=SUM(M31:N31)` | **Invariante** | Totales mensuales de la tabla anual estática; SUM puro; hoja protegida vigente (D2b). |
| S-MN-3 | Refs col M/N | `INTERVENTORIA!O26:O30` (`Mx=Rx`) | **Invariante** | Compara la columna M contra R (ambas Segunda); no depende de la quincena del run. |
| S-MN-4 | Refs col M/N | `INTERVENTORIA!Q26:Q31 = CONSOLIDADO_TOTAL RECAUDO!L104:L109` | **Invariante (en forma)** | Espejo del total L por ASE; sigue transitivamente la selección L9:L13. |
| S-MN-5 | Pares M/N | `GERENTES_*`, `REMUNERACION_*`, `VALIDACION_*`, `ANTICIPOS USUARIOS`, `ANT EXT-REV`, `AJUSTES`, `Informe AFaseo Recaudo 202609-2`, `Valida -Remunera` (p. ej. `U9=R1!M75`+`V9=R1!N75`; `M23=M9-CONSOLIDADO!M104`+`N23=N9-CONSOLIDADO!N104`) | **Invariante** | El template computa AMBAS columnas en paralelo; son mapeos emparejados M+N, no una selección de quincena por el run. |
| S-VAL | VALIDACION_* | `VALIDACION_*!O3=H3-N3`, `N9=SUM(N3:N8)`, `J/M/E/O` vs `REMUNERACION_*` M/N y `D66/D70/N66/N70` | **Invariante** | Consumidores de `N3` (código de período) y mapeos emparejados; barrido de `2026092` en VALIDACION_* → **0 hits** (sin fragmentos con sufijo de período). |
| S-J | DetRetri J | `BCE SC POR FACT.!H3:H7=DetRetri2026092!J9:J13`, `H10=H9=DetRetri2026092!J14` | **Intencional-por-período — Phase 0 SÍ** | 6 de las 11 `<f>` del token de período (§3); el token se reescribe. `DetRetri!J9` en sí es celda VALOR → invariante (dato). |
| S-DV | DetValiRetri | `DetValiRetri!K24:K28=+DetRetri2026092!N32:N36` | **Intencional-por-período — Phase 0 SÍ** | 5 de las 11 `<f>` del token de período (§3). |
| S-DV | DetValiRetri | `DetValiRetri!I16:I21=I9:I14-CONSOLIDADO!Z104:Z109` | **Invariante** | `Z104:Z109` es fórmula shared (`<f t="shared" si="50"/>`, cachés `-L…`): sigue la selección L. |
| S-DV | DetValiRetri | `DetValiRetri!D21=D14-CONSOLIDADO!U109` | **Invariante** | Mapeo a columna U del consolidado. |
| S-DV | DetValiRetri | `DetValiRetri!O28:O29` (refs `INTERVENTORIA!I7/I8/K15..K19/Q19`) | **Invariante** | Valores anuales estáticos de la tabla INTERVENTORIA. |
| S-AJ | Cadenas AJUSTES | `CONSOLIDADO!D85:D89=AJUSTES!D47:D51` (+E..J,N,U..AE); `AJUSTES!D47=D9+D28`, `K47=D47-E47-…`, `M47=K47-L47`, `O47=M47+N47`, `Z47=-L47`, `AE47=SUM(U47:AC47)`, `AG47=AF47=0` | **Invariante** | Cadenas de cálculo sin rama Q1/Q2; `SALDOS POR NOTA` y `RETRIBUCION NEGATIVA` → **0 refs** a M/N/CONSOLIDADO (barrido). `AjustesSfT=0` en Q1 es invariante de CÁLCULO (dominio/fuentes), no de plantilla. |
| S-R26 | INTERVENTORIA R26 | `R26 = J15` (shared `R26:R31`, → `J15:J19`) | **Invariante-en-forma / divergencia cosmética versionada** | Referencia fija (no `F15`/`H15` como citaba el puntero §2.3): `J15:J19` = valores anuales. No re-abierto. Consecuencia: `S26:S31=Q=R` → ver hallazgo NUEVO. |
| S-C59 | Banco C59 | `REPORTE RECAUDO x BANCO!C59` = valor plano `2` (sin `<f>`) | **Invariante de ESCRITURA por dominio** | El runtime la sella por dominio (`OpenXmlPlantillaWriter`: `EscribirValorNumerico(..., "C59", leaf.ReporteBanco.Quincena, ...)`); Phase 0 no la toca. `2` es placeholder del template. |
| S-N3 | Consumidores N3 | `REMUNERACION_*!N3=CONSOLIDADO!N3`; `GERENTES_*!N3`, `GERENTES_TOTAL!N3=CONSOLIDADO!N98`, `AJUSTES!N3=CONSOLIDADO!N3`, `CONSOLIDADO!AD3=N3`, `N22/N41/N60/N79/N98/M117=N3`; `VALIDACION_*!O3=H3-N3`, `N9=SUM(N3:N8)`; `ANT EXT-REV!P3=N3=(-VALIDACION_TOTAL!M9)` | **Invariante** | Todos LEEN `N3`; ninguno lo escribe. Phase 0 lo sella como valor plano (`CodigoCompleto`). |

### 6.2 Hallazgo NUEVO (reportar, NO implementar)

> **F-NEW-1 — Familia de auto-checks booleanas que hardcodean la columna M (Segunda quincena).**
> **Veredicto: intencional-por-quincena NO MANEJADO por Phase 0** → requiere su propio T0/decisión.
> **Evidencia (bytes):** son las ÚNICAS celdas —fuera de `CONSOLIDADO!L9:L13`— que referencian
> `INTERVENTORIA!M26:M31`, y traen EXACTAMENTE la forma M en el oficial Y en las 3 bases preparadas
> offline (incluida la Q1 `2026071`): el preparador offline nunca las conmutaba.
>
> | Celda | `<f>` | Caché | Naturaleza |
> |---|---|---|---|
> | `DetRetri{per}!L32:L36` | `INTERVENTORIA!M26:M30 + L23:L27 = L9:L13` | 1 (TRUE) | Check booleano |
> | `DetRetri{per}!O32:O36` | `(… − INTERVENTORIA!M26:M30) … = O9:O13` | 1 (TRUE) | Check booleano |
> | `DetValiRetri{per}!I24:I29` | `−INTERVENTORIA!M26:M31 + I16:I21 = I9:I14` | 1 (TRUE) | Check booleano |
> | `INTERVENTORIA!S26:S31` | `Q=R` (con `R26:R31 = J15:J19` = Segunda) | 1 (TRUE) | Check booleano |
>
> **Por qué es un hallazgo:** R-INTERV-Q conmuta `CONSOLIDADO!L9:L13` → `N` en Q1, de modo que
> `CONSOLIDADO!L104:L109` (=`L9+…`) pasa a N, mientras estas celdas siguen comparando contra `M`
> (o contra `J15:J19` = M). En una corrida Q1 evaluarían **FALSE** (diferencia material ±1 por ASE
> por el split de redondeo). En la base Q1 retirada NO ocurría porque `L9:L13` también quedaba en M
> (el defecto que R-ATRIBUCION cierra): la inconsistencia es **latente y preexistente**, y Plan 39 la
> expone al corregir L.
>
> **Impacto:** son celdas de **auto-chequeo** (caché TRUE), NO montos de negocio; no alimentan ningún
> total. **No se implementa aquí** (fuera de la allow-list `<f>` de Phase 0, §4): se eleva como
> follow-up con T0 propio. **Sin supuesto silencioso:** queda declarado, no se silencia.

### 6.3 Superficies adicionales barridas (sin hallazgos duros)

- `definedNames`: **0** contienen el token de período (`202609`) → Phase 0 no tiene definedNames que
  reescribir en este template (no-op verificado).
- `docProps/app.xml` y `xl/workbook.xml`: 3 nombres de hoja de período (`Informe AFaseo Recaudo
  202609-2`, `DetRetri2026092`, `DetValiRetri2026092`) → intencional-por-período, manejado (rename).
- `xl/sharedStrings.xml`: `si[340]=2026092`, `si[266]="Segunda quincena"`, `si[267]="Primera
  quincena"` (confirman M=Segunda / N=Primera), `si[333]="QUINCENA"` → el token se reescribe; las
  etiquetas de quincena son rótulos fijos (no se renombran).
- `Valida -Remunera!M/N`, `Z9:Z14=CONSOLIDADO!Z104:Z109` → invariantes (mapeos emparejados).
- `DetValiRetri!O24` ref `INTERVENTORIA!J4` → valor anual estático (invariante).

### 6.4 Regla de cierre

Ningún hallazgo queda en "probablemente igual": S-MN/S-VAL/S-J/S-DV/S-AJ/S-R26/S-C59/S-N3 citados
celda-a-celda; el único ítem sin decisión de negocio es **F-NEW-1** (elevado; no bloquea T4: no toca
el camino L/N3/hojas). Un cambio del oficial por el negocio exige **nuevo T0** (§5).
