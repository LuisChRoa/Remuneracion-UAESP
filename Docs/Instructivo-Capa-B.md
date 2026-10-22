# Instructivo de Capa B — Verificación post-Excel de la Remuneración Quincenal

> **Qué es:** procedimiento ejecutable que el Ingeniero sigue para certificar que el archivo generado por la aplicación reproduce el resultado manual conocido (golden) con tolerancia **±0.5** y fórmulas intactas.
> **Cuándo se usa:** tras cada generación de `Remuneración AAAAMM-# Total.xlsx` (UI o CLI), antes de declarar pase/falla por caso.
> **Rector:** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` §10 (CA-4 fórmulas intactas, CA-6 log por paso) y §10.2 (criterio de éxito).
> **Complemento:** `Docs/Manual-Usuario-Remuneracion-UAESP.md`.

---

## 1. Prerrequisitos

- Windows con **Excel** instalado.
- Los archivos de referencia (goldens) presentes en `Docs/Insumos/` (ver §2 — ancla de integridad).
- La aplicación generada (UI o CLI) y los archivos fuente del período en su carpeta `REMUNERACION AAAAMM{Q}`.

---

## 2. Acta de integridad (SHAs de ancla — NO se re-fijan)

Verifique con `Get-FileHash -Algorithm SHA256` antes de cada sesión de certificación. **Si un SHA difiere del acta, DETÉNGASE y ESCALE** (no "ajuste" el hash; no continúe con un insumo alterado).

| Archivo | SHA256 (acta W-1 HU-16 / HU-17 V2) |
|---|---|
| `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` (golden Q1) | `0B090E9C4851D86DAC534F2965E4A38393CC5BC09350C2A824FD48F7931C096F` |
| `Docs/Insumos/Remuneracion 202607-2 Total.xlsx` (golden Q2) | `584310105AC7223CE840833E3CB64E26F95B61907C9ECD959A0A632EEC21DCB1` |
| `Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx` (canónico Q2) | `95825422B32FBE9E2B60F35C9638E492455FB98CFA182976B657FAC3520E0B8C` |
| `Docs/Insumos/REMUNERACION 2026072/Plantilla  _ Remuneracion 202607-2 Total.xlsx` (control Q2 — NUNCA oráculo) | `509BF210435138B3D36947582E295F27CE1F4D9452BEB1A8D752F3EF4F57529B` |

```powershell
Get-FileHash -Algorithm SHA256 -LiteralPath "Docs\Insumos\Remuneracion 202607-1 Total.xlsx"
```

---

## 3. Protocolo (aplica a cada caso CT)

### Paso 0 — Preparar la salida

Genere la salida en una **carpeta distinta** a la del golden. La salida nunca debe coincidir con la plantilla (la aplicación lo deniega con `ERR-PLANTILLA`, salida 2). El archivo se nombra solo: `Remuneración AAAAMM-# Total.xlsx`.

### Paso 1 — Generar con el CLI

```bash
# Q1 5-ASE (CT-Q1-5A)
Remuneracion.Cli --periodo 2026071 --carpeta "Docs/Insumos/REMUNERACION 2026071" --plantilla "Docs/Insumos/Remuneracion 202607-1 Total.xlsx" --salida "Docs/Insumos/Salidas" --cinco-ase --sobrescribir

# Q1 1-ASE 3 (CT-Q1-1A)
Remuneracion.Cli --periodo 2026071 --carpeta "Docs/Insumos/REMUNERACION 2026071" --plantilla "Docs/Insumos/Remuneracion 202607-1 Total.xlsx" --salida "Docs/Insumos/Salidas" --ase 3

# Q2 5-ASE (CT-Q2-5A)
Remuneracion.Cli --periodo 2026072 --carpeta "Docs/Insumos/REMUNERACION 2026072" --plantilla "Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx" --salida "Docs/Insumos/Salidas" --cinco-ase --sobrescribir
```

Exigencia: línea final `RESULTADO OK codigo=0 ... runId=<guid>` y código de salida **0**. Anote el `runId` en el registro de evidencia.

### Paso 2 — Recalcular en Excel

1. Abra el archivo de salida en Excel.
2. Deje que Excel **recalcule** las fórmulas (Ctrl+Alt+F9 si su configuración no recalcula automáticamente; guardar con Ctrl+S tras el recálculo).
3. **NUNCA** edite fórmulas ni mueva filas; solo lea y compare.

### Paso 3 — Comparar por bloque contra el golden

Compare **celda a celda** el archivo de salida contra el golden del mismo período. Tolerancia: **|diferencia| ≤ 0.5** (redondeo). Registre cada bloque en el registro de evidencia (§7).

---

## 4. Bloque Q1 (golden `Remuneracion 202607-1 Total.xlsx`)

| # | Bloque | Celdas | Referencia golden (202607-1) |
|---|---|---|---|
| Q1.1 | CONSOLIDADO — TOT_OPT | `CONSOLIDADO_TOTAL RECAUDO` D9:D13 | 16704332434.57 · 12157780441.19 · 10101988514.82 · 10552409503.83 · 8519310329.28 |
| Q1.2 | CONSOLIDADO — R2 Total Oportuno | D28:D32 | 54216385.68 · 79400801.26 · 31111803.75 · 17236000.33 · 30774154.23 |
| Q1.3 | CONSOLIDADO — EXTEMP | D47:D51 | 11673020 · 0 · 10198723.07 · 5419780 · 0 |
| Q1.4 | CONSOLIDADO — Reversión R4 | D66:D70 | −12054255.65 · −9889189.72 · −16103442.89 · −21288908.57 · −6421274.68 |
| Q1.5 | CONSOLIDADO — AJUSTES-SF-T | D85:D89 | 0 · 0 · 0 · 0 · 0 (Q1: los ajustes no aplican) |
| Q1.6 | CONSOLIDADO — Total por ASE | D104:D108 | 16758167584.60 · 12227292052.73 · 10127195598.75 · 10553776375.59 · 8543663208.83 |
| Q1.7 | CONSOLIDADO — Gran Total | D109 | 58210094820.50 |
| Q1.8 | DetRetri | `DetRetri2026071` D9:D13 y D14 | Enteros = ROUND(D104:D108,0): 16758167585 · 12227292053 · 10127195599 · 10553776376 · 8543663209 · 58210094822 (D14) |
| Q1.9 | Bloques REMUNERACION_* | `REMUNERACION_ENEL`, `REMUNERACION_OCCIDENTE_ Directa`, `REMUNERACION_ENERBIT`, `REMUNERACION Reciprocidad EAAB`, `REMUNERACION_EAAB-CL` | 100 % fórmulas hacia R1/R2/R4 + CONSOLIDADO; verificar que sigan siendo fórmulas y que sus visibles cuadren con los bloques R1/R2/R4 |
| Q1.10 | BCE | `BCE SC POR FACT.` D3:D7 (CONTRIBUCION), E3:E7 (SUBSIDIO), F3:F7 (TOTAL BSC), H3:H7 (SISTEMA) | D: 3256235169.97 · 1946491828.39 · 1936109874.45 · 474156195.46 · 1266769543.80; E: −1223871491.13 · −6958007481.98 · −3132932285.76 · −438200379.42 · −2477195172.11; F: 2032363678.84 · −5011515653.59 · −1196822411.31 · 35955816.04 · −1210425628.31; H ≈ ROUND(F,0) |
| Q1.11 | VALIDACIONES | `VALIDACION_TOTAL` O9=0, P9=TRUE; `VALIDACION_{EMPRESA}` O=0, P=TRUE por empresa | Oráculo de lectura: O (Recaudo vs REMUNERACION) = 0 y P (INT(O)=0) = TRUE en los bloques activos |
| Q1.12 | INTERVENTORIA | `INTERVENTORIA` K26:K30, M26:M30, N26:N30; K31/M31/N31 SUM; K32=SUM(M31:N31) | Tabla anual (idéntica en Q1 y Q2 — ver §6). K31=1719872614, M31=859936309, N31=859936305, K32=1719872614. Fórmulas intactas; nada escrito por la app |

> Nota Q1.8: el caso single-ASE de referencia (HU-03) usa el total Promoambiental 39148067546.64 → **DetRetri = 39148067547**; en el golden multi-ASE Q1 la misma composición aplica por ASE sobre D104:D108 (valores de la tabla).

---

## 5. Bloque Q2 (golden `Remuneracion 202607-2 Total.xlsx`)

| # | Bloque | Celdas | Referencia golden (202607-2) |
|---|---|---|---|
| Q2.1 | CONSOLIDADO — TOT_OPT | D9:D13 | 17339793559.86 · 20352268540.72 · 15226209207.13 · 7047492494.73 · 12033011685.71 |
| Q2.2 | CONSOLIDADO — R2 Total Oportuno | D28:D32 | 125948553.12 · 178539092.16 · 26586328.29 · 161461949.73 · 42983502.41 |
| Q2.3 | CONSOLIDADO — EXTEMP | D47:D51 | 49580 · 56190070 · 1307680 · 4343101 · 441849 |
| Q2.4 | CONSOLIDADO — Reversión R4 | D66:D70 | −16536713.09 · −71069759.00 · −32310979.80 · −33738104.23 · −3092375.53 |
| Q2.5 | CONSOLIDADO — AJUSTES-SF-T | D85:D89 | **973693.46 · 216025.77 · 104231.83 · 35954.44 · 0** |
| Q2.6 | CONSOLIDADO — Total por ASE | D104:D108 | 17450228673.35 · 20516143969.65 · 15221896467.45 · 7179595395.67 · 12073344661.59 |
| Q2.7 | CONSOLIDADO — Gran Total | D109 | 72441209167.71 |
| Q2.8 | DetRetri-D | `DetRetri2026072` D9:D13 y D14 | Enteros = ROUND(D104:D108,0): 17450228673 · 20516143970 · 15221896467 · 7179595396 · 12073344662 · **72441209168 (D14)** |
| Q2.9 | AJUSTES composición | `SALDOS POR NOTA` + `RETRIBUCION NEGATIVA` (hojas por ASE) → `AJUSTES - SF-T` D47:D51 | Los valores D85:D89 provienen de estas hojas; verificar celdas escritas contra las fuentes (solo 2.ª quincena) |
| Q2.10 | INTERVENTORIA | `INTERVENTORIA` K26:K32 | **Idéntica a Q1** (tabla anual; §6) |

---

## 6. INTERVENTORIA — tabla anual (Q1 y Q2 idéntica)

| ASE | K (Valor oficial mes) | M (2.ª quincena) | N (1.ª quincena) |
|---|---|---|---|
| 1 Promoambiental | 378371975 | 189185988 | 189185987 |
| 2 Lime | 533160511 | 266580256 | 266580255 |
| 3 Ciudad Limpia | 309577071 | 154788536 | 154788535 |
| 4 Bogotá Limpia | 240782166 | 120391083 | 120391083 |
| 5 Área Limpia | 257980891 | 128990446 | 128990445 |
| **Total (K31/M31/N31)** | **1719872614** | **859936309** | **859936305** |
| **Gran total K32** | **1719872614** (= SUM(M31:N31)) | | |

La app **no escribe** este bloque (insumo externo declarado, D2b); la verificación es que la salida lo conserve **intacto** y con las fórmulas SUM.

---

## 7. Fórmulas intactas (CA-4 — spot-check de protegidas)

Verifique que las celdas protegidas del template sigan siendo **fórmulas** en la salida (nunca valores):

| Hoja | Celdas a spot-check |
|---|---|
| `CONSOLIDADO_TOTAL RECAUDO` | D104:D109 (sumas por fila + Gran Total SUM) |
| `AJUSTES - SF-T` | D47:D51 (composición de ajustes) |
| `DetRetri2026071` / `DetRetri2026072` | D23:D28 + D32:D36 (fórmulas) y `DetValiRetri…` equivalentes |
| `REMUNERACION_*` | Todas (100 % fórmulas hacia R1/R2/R4/CONSOLIDADO) |
| `BCE SC POR FACT.` | Filas 9/11 (totales) y columna H (SISTEMA) |
| `INTERVENTORIA` | K31, M31, N31, K32 (SUM) |

Método: en Excel, seleccione la celda y confirme que la barra de fórmulas muestra la fórmula (no un valor); o abra el XML con un editor (dump `<f>`).

---

## 8. Notas operativas

### L1 — Bloque vacío vs fuente vacía (no asumir)

Si un bloque de la fuente está **vacío**, la aplicación **no limpia** el bloque correspondiente de la salida (comportamiento pinneado por test): un bloque en blanco en la salida puede deberse a una fuente vacía y NO a un fallo de escritura. **Verifique la fuente**, no asuma: confirme que la fuente realmente trae ceros/vacío antes de dar por bueno el bloque.

### R1-interior — auditoría fila a fila por sección-contrato + canal efectivo (Planes 45/46)

La hoja `Reporte Componentes R1` se llena por **fila fuente × (fuente, canal, concepto, sección-contrato)**: cada fila de `Recaudoporcomponente` (la que trae el desglose de componentes) cae en el slot de su **sección-contrato** y de su **canal efectivo** (nunca por orden físico). Para auditar:

1. Tome la fila de la fuente `Recaudoporcomponente` (`Sheet1`) y lea su `(A,B,C,D,E)` y los **encabezados de columna** (F..Z).
2. **Sección-contrato = la que CIERRA la sección de la fila** (la frontera `Mes`/`Aplicacion nuevos x reversion` inmediatamente **debajo**, cuyo importe es su subtotal), alineada al oficial de forma monotónica por tipo. La **1.ª sección NO es fija**: si su frontera de cierre es `Aplicacion`, cuenta como Aplicación (y el bloque cae en la `Aplic#1` oficial, no en la 1.ª).
3. **Canal efectivo:** si la fila no trae `D` (`Vlr Servicio`, `Vlr Intereses`, `Subsidio…`), su canal es el `D` de la próxima fila `…/Total` hacia abajo (con `D` no vacío y ≠ `Total`). La fila destino debe tener el **mismo canal efectivo**.
4. Compare **componente a componente por encabezado** (no solo F): `Total`, `Componente TDF`, `TTL`, `TVIAT`, …, `Rec.Intereses` (la columna `Rec.Intereses` de la fuente → **Z** de la plantilla).
5. Busque la fila destino en `Reporte Componentes R1` que tenga el **mismo `(A,B,C,D,E)`** dentro de la sección-contrato/canal correctos. Referencia: `Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md` + su **§12 (ERRATA Plan 46)**.
6. Si el bloque no es **colocable** con el contrato (p. ej. la forma `[M,M,M]` de agosto-ASE3/Q1), el escritor reproduce el comportamiento histórico (neutralidad Q1, plan-then-write). No es un fallo de auditoría.
7. **Reporte de divergencias declaradas:** si el run cierra con `código 6`, junto a la salida está `<salida sin extensión> - Divergencias declaradas.txt` (qué celda, observado vs esperado, Δ, dueño, evidencia, condición de retiro). Ahí se ve qué se reconoció y de quién es la deuda.
8. **Ceros documentados:** hoy **NO hay entradas activas** (el Plan 46 retiró las 2 de `Vlr Intereses` de la 1.ª: ese "hueco" no existe, tienen slot real). Cualquier fila sin slot que no esté en una entrada vigente del catálogo es `ERR-SIN-SLOT` accionable; **no se parchea la plantilla**.

### R2/R4-detalle — auditoría fila a fila por (firma + canal efectivo compuesto) (Plan 47)

Las hojas `Rem. Anticipos R2` y `Reversion Pagos R4` se llenan por **fila fuente × (firma, canal efectivo)**, nunca por orden físico. Para auditar un bloque:

1. **Firma de la fila:** R2 `A|B|C|D`, R4 `A|B|C` (las etiquetas del bloque, en el orden del template).
2. **Canal efectivo (R46-1):** si la fila no trae etiqueta de grupo (R2: `A,B,C` vacíos; R4: `A` vacío), su canal es el rótulo de la fila `Total` que **CIERRA su grupo**; si la fila ES ese cierre (`C=E D=Total` en R2), vale su propio rótulo. Las filas con etiqueta propia fuera del grupo (roles `Componente`/`Subs/Cont`, empresa, `Total` del bloque) se localizan por **firma**.
3. **Canal COMPUESTO (W-3):** en R2 el canal se compone con la **empresa del primer cierre de empresa por debajo** (`B=ENEL C=Total`) → `E/ENEL`, `E/ENERBIT`, `9/NUEVO ESQUEMA`, `A/EAB`… Esto desambigua dos grupos del MISMO canal (ASE5: `E` de ENEL y `E` de ENERBIT) sin recurrir al orden. En R4 no hay composición (su rótulo de grupo YA es la empresa).
4. **Columnas por ENCABEZADO:** R2 `E..R` (`Especiales K` = 0 si la fuente no la trae; `Deb/Cred-otros R` = 0 si no existe), R4 `D..O` + `Otros → R` (13 columnas declaradas).
5. **Ceros documentados del detalle (Plan 48):** `CerosDocumentadosDetalleR2R4` tiene **3 claves / 3 filas** (todas de **agosto-2**, mitad Subs/Cont de BOGOTÁ: `|||Subsidio(-)/Contribucion(+)` / `||A|Total` / `|EAB|Total|`). Las **9 filas de julio-2** que el Plan 47 documentaba (canal `9` y `Vlr Intereses` del canal 2 de LIME/BOGOTÁ) eran **falso hueco**: la v2 les da slot y se escriben. Cualquier otra fila sin slot es `ERR-SIN-SLOT` accionable (hoja + ASE + nº de fila + firma + canal).
6. **Evidencia:** `Docs/T0-Plan48-Adopcion-Plantilla-v2.md` §2.2/§2.4 (falso hueco cerrado) + `Docs/T0-Plan47-Revision-Pestanas-LenteR1.md` §4/§6/§12 + fixtures `RevisionPestanasLenteR1T0Plan47Tests` (matriz v1 histórica + slots v2) y `DetalleR2R4EscrituraTests` (W-2/W-3).

### v2 — base única y control de recaudo (Plan 48)

Desde **2026-10-19** la **v2** (`Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx`, SHA `FF5CF7…2632`) es la **base única** del pipeline; la **v1** (`3086AC31…C5214`) queda **histórica** (solo el caso Q1 del end-to-end la usa como dato del test). Para auditar una salida v2:

1. **Bloques R2 v2:** ASE1 `3-33`, ASE2 `67-124`, ASE3 `175-205`, ASE4 `293-350`, ASE5 `404-441`. El cuadro **OPORTUNO** (filas `B=TOTAL/RECIPROCIDAD/ENEL/ENERBIT/OCCIDENTE/CIUDAD LIMPIA-ACUEDUCTO`, `C=OPORTUNO`) **NO es escribible**: sus filas son **fórmula** (TOTAL/Subsidio/ENEL/OCCIDENTE y RECIPROCIDAD en ASE2/ASE4) o **literal-0** (empresas ausentes). Si una corrida intentara escribir ahí, el guard anti-fórmula lanza `ERR-PLANTILLA`.
2. **Falso hueco cerrado:** verifique `Rem. Anticipos R2` ASE2 `E87/E89 = 831.919,29` y ASE4 `E313/E314/E315 = 20.152.773,11 / -2.377,86 / 20.150.395,25` (antes quedaban en 0; brecha de 53.283.379,41).
3. **Control de recaudo (`Valida - Control Recaudo`):** `D7:D10` se pueblan **por etiqueta** de la columna `C` (ENEL / RECIP., ENERBIT, EAAB / RECIP., DIRECTA OCCIDENTE) desde la fila 27 (TOTAL) de cada `Recaudo *`. En la v2 `E7=SUM(D7:D9)`, `E10=SUM(D10)` y `F10=E7+E10` son **fórmulas**; el writer **sella su `<v>` evaluado** (no las sobrescribe). Spot-check julio-2: `F10 = 72.497.949.967`; agosto-2: `77.064.563.298`. `VALIDACION_TOTAL!C14` debe leer ese valor y `C15 = 1`.
4. **R1 682/702:** la zona de fórmulas que el manual añadió (`F682/F692/F702`, shared `H682:AO682`/`H702:AO702`) **NO se adopta** (delta documentado, P5a pendiente); la v2 no las trae.
5. **`<f>`:** la App no crea ni modifica fórmulas; solo el sellado autorizado escribe `<v>` (nunca `<f>`). Cero `InsertRow`/`DeleteRow`.

### A1 — gate absoluto del EXTEMPORÁNEO (Plan 49; corrige el punto 4 de la sección v2)

Si una corrida termina con un error que contiene **`Plan 49 A1`**, significa que el EXTEMPORÁNEO de un ASE **no llegó al CONSOLIDADO** de la salida. Ejemplo julio-2:

```text
ASE 5 · CONSOLIDADO_TOTAL RECAUDO · D51: EXTEMP resuelto desde lo escrito (0) != Extemp del dominio (441849.00).
Diferencia=441849.00 > ±0.5. (Plan 49 A1: gate absoluto anti-cancelacion.)
ASE 5 · Reporte Componentes R1 · F682: ancla EXTEMP resuelta (0) != suma de secciones F660+F629−L629 (441849.00).
```

1. **Qué pasó:** la plantilla **v2** (heredado de la legacy) trae las anclas `Reporte Componentes R1!F682` (consume `CONSOLIDADO!D51`) y `F702` (consume `REMUNERACION_OCCIDENTE!D51`) como **literal 0 sin fórmula**. La app no las escribe (values-only, fuera de su contrato) y la plantilla no las calcula ⇒ el EXTEMPORÁNEO del **ASE5** no llega al CONSOLIDADO. El gate viejo (`D34/F34/O25`) **no lo veía** porque la pérdida entraba en sus dos operandos y **cancelaba**.
2. **Por qué el punto 4 de la sección v2 queda corregido:** `F682/F702` **no** son "delta del manual": son **anclas consumidoras**. No pueden congelarse como diferencia esperada (regla: ninguna celda referenciada por otra hoja se congela como delta sin demostrar que no altera al consumidor).
3. **Qué hacer:** **NO** editar la plantilla, **NO** forzar el valor, **NO** aceptar como divergencia declarada (es deuda **NO declarada**: falla en duro, por diseño). Reportar a la **propietaria de la plantilla/UAESP** con el ASE, la celda y el Δ que nombra el error. La corrección es una **v3** de la plantilla (restaurar `F682 = F660+F629-L629` y `F702 = F659+F628-L628`) con su T0 de bytes.
4. **Condición de retiro del rojo:** cuando exista la v3 auditada, se re-congela el mapa de anclas con su T0 y las salidas de julio-2/agosto-2 vuelven a verde **por la plantilla**, nunca reverdeciendo tests ni declarando la diferencia.

### M2 — Especiales ausente

Ninguna fuente de saldos-nota trae columna "Especiales": `TieneColumnaEspeciales = false` está fijado por código y el valor ausente se trata como **0**. No es un error: es el comportamiento esperado (5/5 sin columna).

### Q2 en modo 1-ASE

El modo 1-ASE en quincena 2 **falla por diseño** con `ERR-VALIDACION` (salida 1). Para certificar Q2 use **siempre modo 5 ASE** (CT-Q2-5A).

### Fechas de fuente heterogéneas

El rango de fechas de los archivos fuente Q2 NO es homogéneo (p. ej. `RetribuciónNegativa` de ASE4 trae 16072026–31072026). El localizador busca por prefijo, no por fechas: no renombre ni reordene los archivos.

---

## 9. Criterio de pase (Rector §10.2 literal)

> La aplicación se considera exitosa cuando los valores del CONSOLIDADO_TOTAL RECAUDO generados automáticamente **coinciden exactamente** (o con diferencia ≤ ±0.5 por redondeo) con los valores del resultado manual conocido para el mismo período.

**PASA** si, en todos los bloques del caso:
1. Todo cierra dentro de **±0.5** contra el golden (Q1 o Q2 según caso);
2. Las fórmulas protegidas siguen **intactas** (CA-4);
3. El log registró **cada paso** de la ejecución con RunId (CA-6);
4. La plantilla origen no fue mutada (hash intacto).

**FALLA** ante cualquier divergencia fuera de ±0.5, fórmula alterada, o hash divergente. En caso de FALLA, registre **la celda y los dos valores** (salida vs golden) en el registro de evidencia y escale. **Nunca se "ajusta" el golden.**

---

## 10. Registro de evidencia (tabla en blanco)

| Fecha | Quincena (código) | Caso | Comando / ruta | runId | Resultado (PASA/FALLA) | Bloques con divergencia (celda: salida vs golden) | SHA256 salida | Observaciones |
|---|---|---|---|---|---|---|---|---|
| | | | | | | | | |
| | | | | | | | | |
| | | | | | | | | |

Rellene una fila por ejecución certificada. El SHA256 de la salida se obtiene con `Get-FileHash` (registro de que el archivo certificado no cambió después de la revisión).

---

## 11. Fuente de recaudo que no viene en el período (nuevo — Plan 50)

Las fuentes de recaudo (ENEL, OCCIDENTE, NUEVO ESQUEMA, ENERBIT) **pueden no venir** en un período, en
cualquiera de los tres reportes (recaudo por componente, saldos a favor, reversión) y de forma
independiente entre ellos. **No es un error y no hay que hacer nada**: la aplicación deja en **0** las
casillas de esa fuente y registra la ausencia en el archivo de log y en el resultado de la corrida, con
una línea por fuente y reporte, por ejemplo:

```
ASE 2 · R1:NUEVO ESQUEMA
```

**Cómo leerlo en la revisión:**

| Lo que ve | Qué significa | Qué hacer |
|---|---|---|
| Un **0** en las casillas de una fuente, y la fuente aparece en la lista de ausencias | La fuente no vino en el período (caso normal) | Verificar que el cuadro de esa fuente esté en 0 en el archivo y continuar |
| Un **0** sin ninguna línea de ausencia | La fuente sí vino y su valor es 0 | Nada: es un 0 real |
| Un error de **"posible fuente presente no reconocida"** que nombra un ASE, un reporte y una etiqueta | Hay dinero en el reporte que no corresponde a ninguna fuente conocida (por ejemplo, un nombre de fuente cambiado) | **Detener y escalar**: no se corrige en Excel, se reporta al Ingeniero con el ASE y el reporte que indica el mensaje |
| Un error que dice que **no se encontró una fila** de una empresa | La estructura del reporte está incompleta (no es una fuente ausente) | **Detener y escalar**: solicitar el reporte completo |

Un **0 con ausencia registrada es normal**; un **aviso de etiqueta no reconocida es un defecto** que se
escala, nunca se "arregla" a mano en el archivo.
