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
