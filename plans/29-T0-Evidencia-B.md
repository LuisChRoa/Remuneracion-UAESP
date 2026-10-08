# Plan 29 — T0 Evidencia-B (T0b / T0d / T0e / T0f)

> Evidencia de la Fase 0 (bloqueante) del `plans/29 - cierre-brecha-app-vs-manual.md` (§4 T0).
> Retoma el intento previo: **reutiliza** los dumps ya generados en `%TEMP%\opencode\`
> (`t0b_r2_src_all.txt`, `t0b_r4_src.txt`, `t0b_r4_plant.txt`, `t0b_sig_out.txt`, `t0c_out.txt`,
> `t0c2_out.txt`) y agrega solo extractos nuevos mínimos (`n1_out.txt`, `n2_out.txt`, `n3_out.txt`,
> `n4_out.txt`). **Solo evidencia: cero cambios de código productivo o de tests; `Docs/` y la
> plantilla en solo-lectura.** Método: zip + XML BCL (`System.IO.Compression` + `System.Xml`),
> sin ExcelDataReader / OpenXML / Excel / COM. Toda afirmación remite a disco (archivo / hoja / celda).
>
> **Insumos canónicos.** Plantilla `Docs/Plantilla_Remuneracion.xlsx`. App/manual julio
> `Docs/Prueba Julio-2/Resultado/Remuneración 202607-2 Total.xlsx` /
> `Remuneracion 202607-2 Total Administrativo.xlsx`. App/manual agosto
> `Docs/Prueba2/Resultado/Remuneración 202608-2 Total.xlsx` y
> `Docs/Prueba2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx`.
> Fuentes `Docs/Prueba Julio-2/Insumos/` y `Docs/Prueba2/Insumos/`.
>
> **Fecha:** 2026-10-07

---

## 0. Veredictos (tabla única)

| ID | Pregunta | Veredicto | Archivo + celda | Consecuencia por unidad |
|---|---|---|---|---|
| **T0b** | ¿Es congelable la tabla de detalle R2/R4 por ASE en julio y agosto? | **RECORTE R-DESGOSE-DERIVA (parcial).** R2 filas: julio 5/5 == plantilla; agosto 3/5 (LIME y BOGOTA derivan). R2 columnas: derivan en **todos** los ASE (julio E..P sin `Especiales`; agosto E..Q con `Especiales`). R4 filas: julio 5/5; agosto 2/5. | Plantilla `Rem. Anticipos R2` bloques 3..29 / 63..108 / 159..188 / 276..324 / 378..411; `Reversion Pagos R4` bloques 3..15 / 93..105 / 188..200 / 225..243 / 340..352; fuentes `RerpoteDetalleSaldosaFavor*` y `ReversiónPorComponente*` | Unidad R: escribir solo los ASE cuya malla-fuente == malla-plantilla por período; el corrimiento fijo `Especiales K=0` + TCS→L… de D-C queda **refutado** como mapa posicional → exigir mapeo por encabezado (columnas) / label (filas) o recortar. Ver §2 |
| **T0d** | ¿Cada columna de DetRetri C..O / DetValiRetri D..O tiene origen workbook-interno 10/10? | **SÍ: 12/12 DetRetri y 11/12 DetValiRetri.** Origen = `CONSOLIDADO_TOTAL RECAUDO` fila `104+k` (k=ASE−1): D..O 1:1 para DetRetri; U..AE para DetValiRetri. **SALE: `DetValiRetri2026072!J9` (`AJUSTE A LA DECENA`, −19113) sin origen 10/10.** | Manual-julio `DetRetri2026072!D9:O9`, `DetValiRetri2026072!D9:O9`; `CONSOLIDADO_TOTAL RECAUDO!D104:O104`, `U104:AE104`, `J9` | Unidad D: escribir como literal redondeado las columnas trazables (DetRetri D..O; DetValiRetri D..I,K..O); excluir `J` (AJUSTE A LA DECENA) con motivo. `J` de DetRetri (=BCE F3) entra por arrastre Unidad B |
| **T0e** | ¿La plantilla trae vacías-sin-fórmula las celdas de `REPORTE RECAUDO x BANCO`? | **NO.** La plantilla trae **literales-0** (`<v>0</v>`, sin `<f>`) en toda la rejilla de captura C13:H16 (idéntico por bloque). | Plantilla `REPORTE RECAUDO x BANCO!C13:H16`; manual-julio y manual-agosto `C13:H16` | Unidad P: el micro-fix «no escribir 0 donde la plantilla trae vacío» **NO aplica** (no hay celda vacía-sin-fórmula que preservar). La brecha `0`-vs-vacío es el manual **limpiando** celdas (D13:G16 y H16), un borrado ajeno a la escritura de valores. **RECORTE del micro-fix** |
| **T0f** | ¿Qué divergencia pone `VALIDACION_TOTAL!C15` en False post-recalc? | **Sin divergencia de texto de fórmula** (C3:C9, D3:D9, C14 y C15 idénticos app/manual). El False post-recalc es **real, no caché stale**: lo causa el literal `'Valida - Control Recaudo'!F10` (app=0, manual=72497949967) que `C14` (`=C14=C9`) consume; `C9=SUM(C3:C8)` sí recibe valores no-cero del app. | App/manual-julio `VALIDACION_TOTAL!C3:D9`, `C14`, `C15`; `Valida - Control Recaudo!F10` | Unidad P: para R-P-3 (C15=True) hay que poblar `Valida - Control Recaudo!F10` con el valor derivado (`F10 = VALIDACION_TOTAL!C9` = Σ Recaudo), o declarar `J`/`F10` como SALE/manual-oráculo. La fórmula `C15` no se toca |

**Semáforo T0b..T0f:** OK T0d (gate de Unidad D resuelto, con 1 SALE), OK T0e (premisa del micro-fix refutada), OK T0f (culpable localizado), ATENCION T0b (deriva de malla, obliga recorte, no expande alcance).

---

## 1. Evidencia

### 1.1 Dumps reutilizados (sin regenerar)

| Archivo | Contenido | Uso |
|---|---|---|
| `t0b_r2_src_all.txt` | Estructura + label A/B/C/D de `RerpoteDetalleSaldosaFavor` por ASE (julio y agosto) | malla de filas R2 |
| `t0b_r4_src.txt` | Valores A..F de `ReversiónPorComponente` por ASE (julio y agosto) | R4 fuente |
| `t0b_r4_plant.txt` | `Reversion Pagos R4` de la plantilla (D/E/F/P + `D:f`) | R4 destino en ceros |
| `t0b_sig_out.txt` | Firma `A|B|C|D` fuente-vs-plantilla y diff julio↔agosto por ASE | estabilidad R2 |
| `t0c_out.txt`, `t0c2_out.txt` | Cadena `GERENTES_*`/`REMUNERACION_*`/`Reporte Componentes R1`/`CONSOLIDADO`; H2 | T0d / T0f |

### 1.2 Extractos nuevos mínimos (`%TEMP%\opencode\`)

- `n1_out.txt` — hojas de los 5 workbooks; `Rem. Anticipos R2` plantilla (A1:R32) y fuentes ASE1 julio/agosto; `DetRetri`/`DetValiRetri` C..O (manual y app julio); `REPORTE RECAUDO x BANCO`; `VALIDACION_*`.
- `n2_out.txt` — traza de orígenes (tol ±0.5) DetRetri/DetValiRetri fila 9 del manual-julio; `CONSOLIDADO_TOTAL RECAUDO` filas 9 y 104; hoja banco completa; `VALIDACION_TOTAL` + `Valida - Control Recaudo`.
- `n3_out.txt` — `Recaudo ENEL` D3:G26, `CONSOLIDADO` G125:G129 y `Valida - Control Recaudo` D3:F10 en app-julio / manual-julio / plantilla.
- `n4_out.txt` — `REPORTE RECAUDO x BANCO` C12:J17 de la **salida de la app** (julio).

### 1.3 T0b — evidencia

- **Filas de destino EXISTEN en la plantilla en ceros** (confirmado):
  - `Rem. Anticipos R2` E3:R29 (ASE1) = `<v>0</v>` sin `<f>`: `E3=F3=G3=H3=I3=J3=K3=L3=M3=N3=O3=P3=Q3=R3=0`; labels en D (`D3=Vlr Servicio`, `D4=Vlr Intereses`, `D5=Total`, …). Agregados protegidos: `E43=[E17+E28-K17]`, `K43` fórmula.
  - `Reversion Pagos R4` D3:P15 = `<v>0</v>` sin `<f>`; bloques TOTAL OPORTUNO con fórmula (`r73 D:f=D15-P15`, `r375 D:f=D73+D168+D205+D320+D355`).
- **Encabezado de componentes (R2) deriva entre períodos** (nuevo, no estaba en los dumps):
  - Fuente julio fila 4 (ASE1): `E=Total, F=Componente TDF, G=Componente TTL, H=Componente TVIAT, I=Aprovechamiento, J=CCSA Prest.Aprov., K=Componente TCS, L=Componente TLU, M=Componente TBL, N=Componente TRT, O=CCSA Prest. No Aprov., P=Deb/Cred` → **12 columnas (E..P), sin `Especiales`**.
  - Fuente agosto fila 4 (ASE1): idéntico hasta `J`, luego `K=Especiales, L=Componente TCS, M=Componente TLU, N=Componente TBL, O=Componente TRT, P=CCSA Prest. No Aprov., Q=Deb/Cred` → **13 columnas (E..Q), con `Especiales` en K** (valores reales: `K5=144941.68`).
  - Destino plantilla fila 2: `E=Total … J=CCSA Prest.Aprov., K=Especiales, L=TCS, M=TLU, N=TBL, O=TRT, P=CCSA No Aprov., Q=Deb/Cred, R=Deb/Cred-otros` → **14 columnas (E..R)**.
  - Conclusión: el corrimiento del Plan D-C (`Especiales K=0` + TCS→L…) es válido para **julio**; para **agosto** el `K` de la fuente ya es `Especiales` → un mapa posicional fijo rompería. El mapeo robusto es **por nombre de encabezado**.
- **Estabilidad de filas R2** (`t0b_sig_out.txt`): PROMOAMBIENTAL/CIUDAD LIMPIA/AREA LIMPIA `IGUAL`; **LIME 46r→47r** y **BOGOTA LIMPIA 49r→50r** `DIFIEREN`. Plantilla == julio en 5/5; agosto == plantilla en 3/5.
- **Estabilidad de filas R4** (`t0b_r4_src.txt` vs `t0b_r4_plant.txt`): julio 13/13/13/19/13 == plantilla 5/5; agosto 6/13/7/18/13 → solo LIME(2) y AREA(5) coinciden.
- **Correspondencia de valores confirmada** (muestra): fuente julio ASE1 `E5=104754634.94` → manual/plantilla `E3`; `K5(TCS)=8065770.61` → `L3`; `D5=-15794348.21` (R4) → `D3`; `D11=-762167.27` → `D9`.

### 1.4 T0d — evidencia (origen por columna; manual-julio, fila ASE1 = 9)

Traza `Find-Num ±0.5` sobre el manual; el literal de origen cae en la MISMA fila `104+k` de `CONSOLIDADO_TOTAL RECAUDO` (k = ASE−1; ASE1→104), que es el Total por ASE (no el gran total, que es 109).

| DetRetri col (fila 9) | Concepto | Valor manual | Origen workbook-interno |
|---|---|---|---|
| D | RECAUDO TOTAL | 17450228673 | `CONSOLIDADO!D104` (`=D9+D28+D47+D66+D85`) |
| E | RECAUDO POR DISPOSICIÓN FINAL | 907089032 | `CONSOLIDADO!E104` (fórmula shared; `GERENTES_TOTAL!E9 =['…'!E104]`) |
| F | TRATAMIENTO LIXIVIADOS | 511619082 | `CONSOLIDADO!F104` |
| G | COMPONENTE VIAT | 268690519 | `CONSOLIDADO!G104` (`=G9+G28+G47+G66+G85`) |
| H | APROVECHAMIENTO | 3815252257 | `CONSOLIDADO!H104` |
| I | COMERCIALIZACIÓN 11.4% | 134368143 | `CONSOLIDADO!I104` |
| J | BALANCE SUBSIDIO/CONTRIBUCIONES | −3589454666 | `CONSOLIDADO!J9 =['BCE SC POR FACT.'!F3]`; `CONSOLIDADO!J104` |
| K | RECAUDO BASE DE REMUNERACIÓN | 15402664306 | `CONSOLIDADO!K104 =D104-E104-F104-G104-H104-I104-J104` |
| L | COSTO DE LA INTERVENTORÍA | 189185988 | `CONSOLIDADO!L104 =L9+…`; `INTERVENTORIA!F15` |
| M | REMUNERACIÓN FINAL | 15213478318 | `CONSOLIDADO!M104 =K104-L104` |
| N | CONCEPTOS ESPECIALES | 25277316 | `CONSOLIDADO!N104 =N9+…`; `N9 ='Reporte Componentes R1'!L53` |
| O | REMUNERACIÓN FINAL + ESPECIALES | 15238755634 | `CONSOLIDADO!O104 =M104+N104` |

| DetValiRetri col (fila 9) | Concepto | Valor manual | Origen workbook-interno |
|---|---|---|---|
| D | TCS | 1182502863 | `CONSOLIDADO!U104 =U9+…` |
| E | TLU | 1138203111 | `CONSOLIDADO!V104` |
| F | TBL | 5384239175 | `CONSOLIDADO!W104` |
| G | TRT | 2460648865 | `CONSOLIDADO!X104` |
| H | APROVECH 18.6% | 219070049 | `CONSOLIDADO!Y104` |
| I | COSTO DE LA INTERVENTORÍA (−) | −189185988 | `CONSOLIDADO!Z9 =L9*-1`; `Z104` |
| **J** | **AJUSTE A LA DECENA** | **−19113** | **SALE — manual-externo (0 hits en ±0.5 en todo el workbook)** |
| K | CONCEPTOS ESPECIALES | 25277316 | `CONSOLIDADO!N104` (=DetRetri N9) |
| L | BALANCE SUBSIDIO Y CONTRIBUCIÓN BASADO EN RECAUDO | 1428564690 | `CONSOLIDADO!AE104 =AE9+…` |
| M | DIF BSCF − BSCR | 5018019355 | `CONSOLIDADO!AD104 =-J104+AE104` |
| N | REMUNERACIÓN FINAL + ESPECIALES | 15238755634 | `CONSOLIDADO!O104` (=DetRetri O9) |
| O | VALIDACIÓN RETRIBUCIÓN | 0 | celda-bandera (DetValiRetri O9=0; sin desglose) |

- Confirmado por lectura de `OpenXmlPlantillaWriter` (`EscribirCeldasDetRetriQ2`, L1077-1095): la app escribe **solo `DetRetri D9:D14`** (`WorkbookLeafCellMapQ2.ObtenerDetRetriDestino`/`DetRetriTotal`). El resto (C, E..O) queda en 0 en la salida de la app (app-julio `DetRetri E9:O9=0`).
- Los orígenes `E104/F104/I104/J104/V104..Z104` son **fórmulas shared** (tienen `<f>` sin texto = follower) — no literales; el volcado `n2_out.txt` los marca `F=[]`.

### 1.5 T0e — evidencia

Plantilla `REPORTE RECAUDO x BANCO` (bloque ASE1):
- `C13:H16` = **literal 0** (`<v>0</v>` sin `<f>`): `C13=…=H13=0`, `…`, `C16=…=H16=0`.
- Fórmulas de la rejilla: `I13=[SUM(C13:H13)]`, `I14=[SUM(C14:H14)]`, `C17=[SUM(C13:C16)]`, `H17=[SUM(H13:H16)]`; followers `E17/F17/G17/I15/I16`; `J13=[SUM(C13:H13)=I13]`. `C59=2` (quincena) literal; `C62=[IF($C$59=1,…)]`.

Manual-julio (bloque ASE1): `C13=16577115642.75`, `H13=776222233.37`, `C14=83572548.25`, `H14=17358483.63`, `C15=0`, `H15=11733000`, `C16=0`; **`H16` y `D13:G16` ausentes (vacías)**.
Manual-agosto (bloque ASE1): `C13=17445127414`, `H13=878866230`, `C14=40340694`, `H14=20016520.48`, `C15=0`, `H15=11702167.52`, `C16=0`; **`H16` y `D13:G16` ausentes**.
Salida de la app (julio, `n4_out.txt`): `C13=16577115642.75`, `H13=776222233.37`, `C14=83572548.25`, `H14=17358483.63`, `C15=0`, `H15=11733000`, `C16=0`, **`H16=0` y `D13:G16=0`** (heredados de la plantilla).

Conclusión: la app escribe SOLO `C13:C16` y `H13:H16` (mapa `WorkbookLeafCellMapReporteBanco.EditablesPorAse[1] = ENEL→C13:C16, OCCIDENTE→H13:H16`); el aparente «0-vs-vacío» en D13:G16/H16 es **0 de plantilla** contra **vacío del manual** (el manual borró celdas). No hay celda vacía-sin-fórmula en la plantilla.

### 1.6 T0f — evidencia

`VALIDACION_TOTAL` (texto de fórmula, idéntico app vs manual):
- `C3:C7` = `='Recaudo EAAB Reciprocidad'!F21 + … + 'Recaudo ENERBIT'!F21` (… F22..F25); `C8` análogo con F26; `C9=[SUM(C3:C8)]`.
- `D3:D7` = `=['CONSOLIDADO_TOTAL RECAUDO'!G125]` … `G129`; `D8` literal 0; `D9=[SUM(D3:D8)]`.
- `C14 =['Valida - Control Recaudo'!F10]`; `C15 =[C14=C9]` (caché 1 en ambos).
- `VALIDACION_*` ×5: `C8/C9/C14/C15` con textos idénticos app vs manual.

Diferencia app vs manual: `Valida - Control Recaudo!F10` es **LITERAL** en ambos (app=0, manual=72497949967); los vecinos `D7:F10` del manual son literales (68144618192 / 68213380915 / 28645323 / 40117400 / 4284569052 / 72497949967) y en la app quedan 0. `n3_out.txt` prueba que la app SÍ trae `Recaudo ENEL!F21..F25 ≠ 0`, por lo que `C9` post-recalc ≠ 0 ⇒ `C15 = (0 = C9) = False`. El manual hace `F10 = C9` (72497949967) ⇒ True.

`WorkbookLeafCellMapValidaciones` (L142-143) confirma que `Valida - Control Recaudo` es **protegida-valor** («mezcla VALORES (F10) con booleanos (F7) → no exigir fórmula») y que `C15` está en `SubBloquesValidacionTotal` (gate TRUE exacto).

---

## 2. Tabla congelada T0b (fuente → destino)

### 2.1 R2 — columnas (resolución por ENCABEZADO, no posicional)

| Fuente julio (fila 4) | Fuente agosto (fila 4) | Destino plantilla (fila 2) | Regla |
|---|---|---|---|
| E = Total | E = Total | E = Total | nombre |
| F = Componente TDF | F = Componente TDF | F = Componente TDF | nombre |
| G = Componente TTL | G = Componente TTL | G = Componente TTL | nombre |
| H = Componente TVIAT | H = Componente TVIAT | H = Componente TVIAT | nombre |
| I = Aprovechamiento | I = Aprovechamiento | I = Aprovechamiento | nombre |
| J = CCSA Prest.Aprov. | J = CCSA Prest.Aprov. | J = CCSA Prest.Aprov. | nombre |
| (ausente) | K = Especiales | K = Especiales | julio: 0 (invariante) |
| K = Componente TCS | L = Componente TCS | L = Componente TCS | nombre |
| L = TLU | M = TLU | M = TLU | nombre |
| M = TBL | N = TBL | N = TBL | nombre |
| N = TRT | O = TRT | O = TRT | nombre |
| O = CCSA Prest. No Aprov. | P = CCSA Prest. No Aprov. | P = CCSA Prest. No Aprov. | nombre |
| P = Deb/Cred | Q = Deb/Cred | Q = Deb/Cred | nombre |
| (ausente) | (ausente) | R = Deb/Cred-otros | 0 (sin contraparte) |

> El «corrimiento» `TCS→L/TLU→M/TBL→N/TRT→O/CCSA-NoAprov→P/DebCred→Q` es la proyección **julio**; en agosto la fuente ya trae `Especiales` en K y el mismo mapeo por nombre sigue aplicando.

### 2.2 R2 — filas / bloques por ASE

| ASE | Plantilla nameRow | Detalle destino | Fuente julio | Fuente agosto | jul==plantilla | ago==plantilla |
|---|---|---:|---:|---:|:--:|:--:|
| 1 PROMOAMBIENTAL | 1 | 3..29 (27) | 5..31 (27) | 5..31 (27) | MATCH | MATCH |
| 2 LIME | 61 | 63..108 (46) | 5..50 (46) | 5..51 (47) | MATCH | **NO-MATCH** |
| 3 CIUDAD LIMPIA | 157 | 159..188 (30) | 5..34 (30) | 5..34 (30) | MATCH | MATCH |
| 4 BOGOTA LIMPIA | 274 | 276..324 (49) | 5..53 (49) | 5..54 (50) | MATCH | **NO-MATCH** |
| 5 AREA LIMPIA | 376 | 378..411 (34) | 5..38 (34) | 5..38 (34) | MATCH | MATCH |

Offset constante por bloque: `destino = fuente − 2` (ASE1: fuente 5→destino 3; … fuente 31→destino 29).
**Filas exactas que cambian en agosto** (`t0b_sig_out.txt`):
- **LIME (+1 fila neta, 46→47):** primera divergencia en el índice 8 del bloque — julio=`Vlr Intereses` / agosto=`2|Total`; agosto intercala el sub-bloque componente `2` (`C=2|D=Total`) y desplaza el resto +1 (cierre `Total` pasa de `[45]` a `[46]`).
- **BOGOTA LIMPIA (+1 fila neta, 49→50):** primera divergencia en el índice 1 — julio=`Vlr Intereses` / agosto=`A|Total`; agosto intercala el bloque **EAB** (`C=A|D=Total` + `B=EAB|C=Total`) al inicio y lo repite en la mitad Subs/Cont (tras índice 28), desplazando +1.

### 2.3 R4 — filas / bloques por ASE

| ASE | Plantilla nameRow | Detalle destino | Fuente julio | Fuente agosto | jul==plantilla | ago==plantilla |
|---|---|---:|---:|---:|:--:|:--:|
| 1 PROMOAMBIENTAL | 1 | 3..15 (13) | 5..17 (13) | 5..10 (6) | MATCH | **NO-MATCH** |
| 2 LIME | 91 | 93..105 (13) | 5..17 (13) | 5..17 (13) | MATCH | MATCH |
| 3 CIUDAD LIMPIA | 186 | 188..200 (13) | 5..17 (13) | 5..11 (7) | MATCH | **NO-MATCH** |
| 4 BOGOTA LIMPIA | 223 | 225..243 (19) | 5..23 (19) | 5..22 (18) | MATCH | **NO-MATCH** |
| 5 AREA LIMPIA | 338 | 340..352 (13) | 5..17 (13) | 5..17 (13) | MATCH | MATCH |

Mapeo de columnas R4-detalle: `D←D(Total)`, `E←E(Componente TDF)`, `F←F(Componente TTL)`, `P(destino)=0` (`SERVICIO ESPECIALES`, sin contraparte en la fuente).
Offset: `destino = fuente − 2`. Filas de cierre con fórmula en plantilla (no escribir): `r73=D15-P15`, `r168`, `r205`, `r320`, `r355`, y generales `r375=D73+D168+D205+D320+D355` / `r376=D74+D169+D206+D321+D356`.

### 2.4 `RECORTE R-DESGOSE-DERIVA` (recorte exacto)

La malla de detalle **no es congelable posicionalmente** en 12 meses:
- **Columnas R2:** derivan en **todos** los ASE (julio E..P / agosto E..Q). Recorte: el mapa debe resolver por **encabezado**; prohibido el corrimiento fijo de D-C para agosto.
- **Filas R2:** deriva en **LIME-agosto** y **BOGOTA-agosto**. Recorte: excluir esos 2 combos del detalle R2 (o resolver por label A/B/C/D).
- **Filas R4:** deriva en **PROMOAMBIENTAL-agosto (13→6)**, **CIUDAD LIMPIA-agosto (13→7)** y **BOGOTA LIMPIA-agosto (19→18)**. Recorte: excluir esos 3 combos del detalle R4.

**Recorte propuesto:** Unidad R se reduce a los combos estables salvo aprobación explícita del mapeo por encabezado/label:
- R2 detalle: julio `{1,2,3,4,5}`; agosto `{1,3,5}`.
- R4 detalle: julio `{1,2,3,4,5}`; agosto `{2,5}`.
- Si el Ingeniero NO acepta el mapeo dinámico, Unidad R se recorta a **«agregados + documentar»** (los agregados R2/R4 ya cierran por V2: `E43=E17+E28-K17`, `D73=D15-P15`). Prohibido insertar/borrar filas (D-C).

---

## 3. Reproducibilidad

Scripts de medición en `%TEMP%\opencode\` (no versionados): `T0lib.ps1`/`t0help.ps1` (helpers zip+XML), `t0b_sig.ps1`, `t0b_r2_src_all.ps1`, `t0b_r4_src.ps1`, `t0b_r4_plant.ps1`, y los nuevos `n1.ps1`, `n2.ps1`, `n3.ps1`, `n4.ps1` (salidas `n*_out.txt`).

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\n1.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\n2.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\n3.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\opencode\n4.ps1"
```

Sin `NEEDS_CONTEXT`: los cuatro veredictos cierran con evidencia en disco. El único recorte (R-DESGOSE-DERIVA) es el previsto por el plan como salida honesta de T0b; no expande alcance.

---

## 4. Anexo — orígenes DetRetri por ASE (fila `104+k`)

Para ASE `n` (DetRetri fila `8+n`, DetValiRetri fila `8+n`) el origen es la MISMA columna en `CONSOLIDADO_TOTAL RECAUDO` fila `104+(n-1)`. Verificación cruzada julio: `DetRetri D9=17450228673` (ASE1) / `D10=20516143970` (ASE2) / `D11=15221896467` (ASE3) / `D12=7179595396` (ASE4) / `D13=12073344662` (ASE5) ↔ `CONSOLIDADO D104..D108`. Fila `14` (Total) ↔ `CONSOLIDADO D109`.
