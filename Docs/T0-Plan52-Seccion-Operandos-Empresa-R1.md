# T0 — Plan 52: mapeo por sección de los operandos empresa R1

| Campo | Valor |
|---|---|
| Plan | `plans/52 - mapeo-por-seccion-operandos-empresa-r1.md` (**APROBADO 2026-10-22**, §12 ratificado 3/3) |
| Fase | **H0 (read-only, bloqueante)** — gate `H0 → H1` |
| Fecha | 2026-10-22 |
| Veredicto | **ROJO — STOP. H1 NO se abre.** (`S1.5` falla y `S1.7` deja un caso no arbitrable; ver §9) |
| Alcance ejecutado | Cero código de producción. Solo lectura de las 20 fuentes R1 reales (`Docs/Prueba {Julio-1, Julio-2, Agosto-2, Octubre-1}/Insumos/{1-Promoambiental…5-Área Limpia}/Recaudoporcomponente_*.xlsx`) + lectura del oráculo manual (`REMUNERACION_*!D9:D13`, `CONSOLIDADO_TOTAL RECAUDO!D47:D51`) + la plantilla oficial v2 (solo lectura) |
| Invariantes verificadas antes/después | v2 `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` · v1 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` (intactas) |

---

## 0. Resumen ejecutivo

El T0 se ejecutó completo (S1.1–S1.7) sobre los 4 períodos × 5 ASE. Resultado por gate:

| Gate | Veredicto | Evidencia |
|---|---|---|
| **S1.1** matriz ASE × empresa × período, firma A–E, sección por cierre y por apertura | **VERDE (con corrección de premisa D1)** | §2: 52 grupos, 180 filas-empresa, todas con firma `\|<label>\|Total\|` (E vacío) |
| **S1.6** tipabilidad por `SeccionPorCierre` | **VERDE** | 0 filas no tipables, 0 empates en la misma sección, 0 secciones intercaladas (§7) |
| **S1.2** ordinal vs magnitud | **ROJO** | 2 grupos verdes con **magnitud ≠ ordinal** (§3) |
| **S1.5** identidad exacta en todo lo verde | **ROJO** | **17 operandos-celda** divergen; 7 de ellos mueven el `VisibleR1` (§6) |
| **S1.4** inventario R2/R4 (D6) | **VERDE** | R2 nunca tiene 1 fila-empresa → el defecto gemelo **no** tiene caso real; R4 nunca tiene >1 fila (§5) |
| **S1.7** arbitraje del oráculo manual | **PARCIAL** | 4 pares arbitrados **a favor de la regla por sección**; `TotalMedio` **no** es observable en el manual (§8) |

**Decisión de H0:** no se abre H1. Los motivos son de contrato, no de implementación:

1. **La lista cerrada §9 está incompleta.** El T0 mide **17** operandos-celda con Δ (no 3 clases de Δ): 2 secciones `TotalMedio→0` (§9-ii), 1 caso `oct-1 ASE3` throw→valores (§9-i) y **13 Δ más** que §9 no enumera (7 de ellos mueven el `VisibleR1`, hasta −12,8 mil millones). §9 dice literalmente: *"Cualquier otro Δ, alta o baja = STOP"*.
2. **El invariante D4 ("identidad bit-exacta en todo lo verde") es falso por datos, no por implementación.** Los duplicados por ocurrencia están hoy en valores "verdes" porque la conciliación por empresa **no se ejecuta en Q2** (recorte HU-20/G2-D1). El oráculo manual confirma que **los valores actuales son el defecto** y que la resolución por sección los corrige al centavo (§8). Es decir: D4 debe re-acotarse a **Q1** (donde la conciliación sí corre) antes de poder usarse como blocker.
3. **La regla D3 ("única fila en primera sección ⇒ `Menor`, `Main = 0`") rompe un ASE completo en octubre-1.** En `octubre-1 ASE5` las 3 empresas tienen **una sola** fila-empresa y está en la **primera** sección (que además es la única sección de recaudo de la hoja): la regla manda `TotalMain = 0`. Para ASE5-ENERBIT el mapa `(3,5)` **no declara `TotalMenor`** ⇒ el importe de 160.186.886,97 **desaparece** y `VisibleR1` queda **negativo** (−20.676.306,97). Ninguna decisión ratificada cubre este caso.
4. **La premisa de D1 ("cierre y apertura coinciden por construcción") es falsa.** En las 180 filas-empresa `SeccionPorCierre` y `PorFronteraArriba` devuelven **etiquetas distintas** siempre (la sección se abre por una frontera y se cierra por otra). La *partición* y el *orden* sí coinciden (lo que la resolución necesita), pero el texto de D1 debe corregirse citando Plan 46 §2.2 (DEFECTO 2) — no es un STOP funcional, es una corrección de justificación.

Los puntos 2/3/4 tienen propuesta de enmienda en §9.2; **la decisión es del Ingeniero** (no se auto-aprueba).

---

## 1. Método (reproducible, read-only)

1. **Secuencia tipada**: `ExcelDataReaderWorkbookLeafInputReader.LeerEspejoR1(ase, ruta)` (producción, sin cambios) construye la lista `FilaEspejoR1` de la hoja `Sheet1` desde la primera fila `Vlr Servicio` hasta la fila `Total` final.
2. **Alineación espejo ↔ fila real**: se leyó la hoja por bytes (`ZipArchive`+`XDocument`, BCL) con el número de fila real de cada `<row>`, se aplicó la misma ventana (primera `Vlr Servicio` → `Total` con B vacío) y se **asertó 1:1 la firma `A|B|C|D|E`** de las 20 fuentes → **20/20 alineadas** (por eso todas las evidencias de este documento citan `archivo:Sheet1:fila` real).
3. **Tipado de sección**: `ContratoSeccionPeriodFreeR1.SeccionPorCierre(filas, i)` (Plan 46, vigente) y `SeccionEspejoR1.PorFronteraArriba(filas, i)` (Plan 41/45, se usa solo como contraste).
4. **Filas-empresa**: misma regla del reader — `C = label(empresa, ase) ∧ D = 'Total'`, separando `indices < indiceComponente` (bloque Totales) de `indices > indiceComponente` (bloque Subs).
5. **Resolución "vieja" (la del código actual)**: `OrderByDescending(|F|)` → `Main = [0]`, `Medio = [1]`, `Menor = [count-1]`.
6. **Resolución "nueva" (la del plan, D2/D3)**: secciones presentes en orden de aparición; `Main` = última, `Menor` = primera, `Medio` = la intermedia (solo con ≥3); con **una sola** sección: si es la primera del bloque → `Menor`, si no → `Main`; ausente → `0`.
7. **Oráculo**: `REMUNERACION_<empresa>!D{8+ase}` del manual del administrativo (julio-2 / agosto-2) y `CONSOLIDADO_TOTAL RECAUDO!D47:D51`.
8. Sin escrituras: **no** se abrió ninguna plantilla en modo escritura, no se generaron salidas, no se tocó el writer.

Ninguna de las filas-empresa trae `E` no vacío (R3 no se dispara) y ninguna sección tiene 2 filas-empresa de la misma empresa (R2/empate no se dispara); ver §7.

---

## 2. S1.1 — Matriz (4 períodos × 5 ASE × empresa del mapa)

`f<n>` = fila real de `Sheet1` de la fuente `Recaudoporcomponente_*.xlsx` del ASE. `Oportuno#k` / `Aplicacion#k` = `SeccionPorCierre`. Columna **Orden** compara la asignación posicional (regla nueva) contra la de magnitud (regla actual): `Main ok` = la última sección es el mayor `|F|`; `Menor ok` = la primera sección es el menor `|F|`; `Medio ok` = la sección intermedia es el `|F|` de índice 1.

| Período | ASE | Empresa | Totales (fila · sección · F) | Subs | Orden posición vs magnitud |
|---|---|---|---|---|---|
| Julio-1 | 1 | 2 `ENEL` | f13 Oportuno#1 = 19,104,244,534.30 | f32 Oportuno#2 = -2,909,727,085.30 | — ‖ — |
| Julio-1 | 1 | 4 `OCCIDENTE` | f8 Aplicacion#1 = 6,331,515.37 · f23 Oportuno#1 = 445,542,416.32 | f28 Aplicacion#2 = 5,341,504.63 · f39 Oportuno#2 = 82,466,308.49 | Main ok / Menor ok ‖ Main ok / Menor ok |
| Julio-1 | 2 | 1 `NUEVO ESQUEMA` | f10 Oportuno#1 = -16,473.56 · f26 Oportuno#2 = 18,903,236.32 | f46 Oportuno#3 = -5,316,702.76 | Main ok / Menor ok ‖ — |
| Julio-1 | 2 | 2 `ENEL` | f7 Oportuno#1 = -1,046,185.37 · f19 Oportuno#2 = 16,423,078,954.74 | f41 Oportuno#3 = -4,636,257,580.37 | Main ok / Menor ok ‖ — |
| Julio-1 | 2 | 4 `OCCIDENTE` | f13 Oportuno#1 = -57,062.45 · f36 Oportuno#2 = 411,038,040.91 | f53 Oportuno#3 = -20,808,604.46 | Main ok / Menor ok ‖ — |
| Julio-1 | 3 | 2 `ENEL` | f7 Oportuno#1 = -11,299.64 · f21 Oportuno#2 = 11,228,997,120.99 | f40 Oportuno#3 = -1,564,808,872.35 | Main ok / Menor ok ‖ — |
| Julio-1 | 3 | 4 `OCCIDENTE` | f10 Oportuno#1 = -9,583.17 · f16 Aplicacion#1 = 5,698,824.85 · f31 Oportuno#2 = 312,921,279.71 | f36 Aplicacion#2 = 4,499,898.22 · f47 Oportuno#3 = 157,220,099.46 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Julio-1 | 4 | 1 `NUEVO ESQUEMA` | f34 Oportuno#2 = 6,598,912.91 | f63 Oportuno#3 = 666,067.09 | — ‖ — |
| Julio-1 | 4 | 2 `ENEL` | f7 Oportuno#1 = -68,295.75 · f16 Aplicacion#1 = 3,576,398.94 · f27 Oportuno#2 = 10,929,648,714.84 | f49 Aplicacion#2 = -12,248.94 · f58 Oportuno#3 = -588,681,732.09 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Julio-1 | 4 | 4 `OCCIDENTE` | f10 Oportuno#1 = -126,294.09 · f22 Aplicacion#1 = 1,713,298.93 · f44 Oportuno#2 = 213,621,752.53 | f54 Aplicacion#2 = 142,331.07 · f70 Oportuno#3 = 13,718,891.96 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Julio-1 | 5 | 2 `ENEL` | f7 Oportuno#1 = -6,448.05 · f16 Oportuno#2 = 8,449,843,621.14 | f35 Oportuno#3 = -209,268,624.09 | Main ok / Menor ok ‖ — |
| Julio-1 | 5 | 3 `ENERBIT` | f20 Oportuno#2 = 147,224,121.37 | f38 Oportuno#3 = -20,851,206.37 | — ‖ — |
| Julio-1 | 5 | 4 `OCCIDENTE` | f10 Oportuno#1 = -6,194.17 · f30 Oportuno#2 = 134,085,436.94 | f45 Oportuno#3 = 45,223,319.23 | Main ok / Menor ok ‖ — |
| Julio-2 | 1 | 2 `ENEL` | f7 Oportuno#1 = -7,108.82 · f15 Aplicacion#1 = 9,215.69 · f20 Oportuno#2 = 15,315,999,861.16 | f35 Aplicacion#2 = 40,364.31 · f39 Oportuno#3 = 1,261,122,890.41 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Julio-2 | 1 | 4 `OCCIDENTE` | f10 Oportuno#1 = -4,108.55 · f30 Oportuno#2 = 619,934,942.23 | f46 Oportuno#3 = 168,024,399.69 | Main ok / Menor ok ‖ — |
| Julio-2 | 2 | 1 `NUEVO ESQUEMA` | f38 Oportuno#2 = 15,303,333.19 | f69 Oportuno#3 = -3,892,093.19 | — ‖ — |
| Julio-2 | 2 | 2 `ENEL` | f7 Oportuno#1 = -381,705.66 · f16 Aplicacion#1 = 30,144,259.37 · f28 Oportuno#2 = 24,064,820,423.33 | f53 Aplicacion#2 = 25,934,600.63 · f62 Oportuno#3 = -5,016,640,264.67 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Julio-2 | 2 | 4 `OCCIDENTE` | f10 Oportuno#1 = -24,201.09 · f23 Aplicacion#1 = 249,172.26 · f48 Oportuno#2 = 1,331,532,739.24 | f58 Aplicacion#2 = -137,962.26 · f76 Oportuno#3 = 15,004,530.85 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Julio-2 | 3 | 2 `ENEL` | f13 Aplicacion#1 = 680,538.03 · f18 Oportuno#2 = 15,107,063,589.67 | f33 Aplicacion#2 = 627,141.97 · f37 Oportuno#3 = -1,096,861,177.67 | Main ok / Menor ok ‖ Main ok / Menor ok |
| Julio-2 | 3 | 4 `OCCIDENTE` | f7 Oportuno#1 = -16,458.46 · f28 Oportuno#2 = 825,593,716.53 | f44 Oportuno#3 = 404,748,758.01 | Main ok / Menor ok ‖ — |
| Julio-2 | 4 | 1 `NUEVO ESQUEMA` | f19 Aplicacion#1 = 41,900.00 · f38 Oportuno#2 = 5,353,948.68 | f67 Oportuno#3 = 262,901.32 | Main ok / Menor ok ‖ — |
| Julio-2 | 4 | 2 `ENEL` | f7 Oportuno#1 = -280,784.65 · f16 Aplicacion#1 = 3,836,285.07 · f28 Oportuno#2 = 7,238,088,945.92 | f53 Aplicacion#2 = -226,278.07 · f60 Oportuno#3 = -633,811,893.27 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Julio-2 | 4 | 4 `OCCIDENTE` | f10 Oportuno#1 = -159,171.87 · f23 Aplicacion#1 = 722,390.09 · f48 Oportuno#2 = 446,469,722.97 | f56 Aplicacion#2 = -31,196.09 · f74 Oportuno#3 = 25,697,170.56 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Julio-2 | 5 | 2 `ENEL` | f18 Oportuno#1 = 12,823,196,468.26 | f43 Oportuno#2 = -1,113,041,546.26 | — ‖ — |
| Julio-2 | 5 | 3 `ENERBIT` | f22 Oportuno#1 = 31,835,017.56 | f46 Oportuno#2 = -4,970,797.56 | — ‖ — |
| Julio-2 | 5 | 4 `OCCIDENTE` | f13 Aplicacion#1 = 417,519.62 · f32 Oportuno#1 = 229,202,451.62 | f39 Aplicacion#2 = 24,329.38 · f53 Oportuno#2 = 81,691,821.42 | Main ok / Menor ok ‖ Main ok / Menor ok |
| Agosto-2 | 1 | 2 `ENEL` | f12 Aplicacion#1 = 42,369.42 · f17 Oportuno#2 = 17,293,233,184.84 | f32 Aplicacion#2 = 21,200.58 · f36 Oportuno#3 = 151,894,229.16 | Main ok / Menor ok ‖ Main ok / Menor ok |
| Agosto-2 | 1 | 4 `OCCIDENTE` | f7 Oportuno#1 = -9,123.90 · f27 Oportuno#2 = 684,286,720.04 | f43 Oportuno#3 = 206,290,801.38 | Main ok / Menor ok ‖ — |
| Agosto-2 | 2 | 1 `NUEVO ESQUEMA` | f34 Oportuno#2 = 11,485,037.33 | f60 Oportuno#3 = -2,134,827.33 | — ‖ — |
| Agosto-2 | 2 | 2 `ENEL` | f7 Oportuno#1 = -1,414,859.14 · f17 Aplicacion#1 = 131,543.19 · f25 Oportuno#2 = 24,836,019,402.94 | f49 Aplicacion#2 = 65,196.81 · f53 Oportuno#3 = -5,574,513,313.02 | Main ok / **Menor≠** (mag f17) / **Medio≠** (mag f7) ‖ Main ok / Menor ok |
| Agosto-2 | 2 | 4 `OCCIDENTE` | f12 Oportuno#1 = -28,650.84 · f20 Aplicacion#1 = 24,890.00 · f44 Oportuno#2 = 2,144,338,413.37 | f67 Oportuno#3 = 116,213,032.47 | Main ok / **Menor≠** (mag f20) / **Medio≠** (mag f12) ‖ — |
| Agosto-2 | 3 | 2 `ENEL` | f7 Oportuno#1 = -3,933.31 · f16 Oportuno#2 = 16,212,894,162.40 | f31 Oportuno#3 = -1,087,996,443.09 | Main ok / Menor ok ‖ — |
| Agosto-2 | 3 | 4 `OCCIDENTE` | f10 Oportuno#1 = -522.80 · f26 Oportuno#2 = 897,711,004.34 | f38 Oportuno#3 = 360,488,591.23 | Main ok / Menor ok ‖ — |
| Agosto-2 | 4 | 1 `NUEVO ESQUEMA` | f22 Aplicacion#1 = 210,280.87 · f41 Oportuno#2 = 4,068,782.15 | f59 Aplicacion#2 = 22,589.13 · f73 Oportuno#3 = 294,267.85 | Main ok / Menor ok ‖ Main ok / Menor ok |
| Agosto-2 | 4 | 2 `ENEL` | f7 Oportuno#1 = -18,514.21 · f16 Aplicacion#1 = 3,718,594.57 · f31 Oportuno#2 = 8,216,147,844.15 | f56 Aplicacion#2 = -75,886.57 · f66 Oportuno#3 = -857,738,903.94 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Agosto-2 | 4 | 4 `OCCIDENTE` | f10 Oportuno#1 = -33,199.73 · f26 Aplicacion#1 = 225,375.38 · f51 Oportuno#2 = 818,884,043.29 | f62 Aplicacion#2 = -70,543.38 · f80 Oportuno#3 = 101,901,939.98 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Agosto-2 | 5 | 2 `ENEL` | f24 Oportuno#2 = 13,260,342,962.39 | f51 Oportuno#3 = -1,503,008,743.39 | — ‖ — |
| Agosto-2 | 5 | 3 `ENERBIT` | f28 Oportuno#2 = 33,465,885.10 | f54 Oportuno#3 = -4,681,675.10 | — ‖ — |
| Agosto-2 | 5 | 4 `OCCIDENTE` | f7 Oportuno#1 = -4,485.01 · f19 Aplicacion#1 = 1,666,506.54 · f38 Oportuno#2 = 252,982,826.29 | f47 Aplicacion#2 = -416,696.54 · f61 Oportuno#3 = 76,680,594.72 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Octubre-1 | 1 | 2 `ENEL` | f7 Oportuno#1 = -2,015.07 · f13 Aplicacion#1 = 546,059.64 · f22 Oportuno#2 = 17,734,947,270.61 | f37 Aplicacion#2 = -32,099.64 · f44 Oportuno#3 = -3,253,607,640.54 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Octubre-1 | 1 | 4 `OCCIDENTE` | f17 Aplicacion#1 = 89,987.43 · f32 Oportuno#2 = 290,244,120.78 | f40 Aplicacion#2 = 44,922.57 · f51 Oportuno#3 = 63,371,932.00 | Main ok / Menor ok ‖ Main ok / Menor ok |
| Octubre-1 | 2 | 1 `NUEVO ESQUEMA` | **ausente** | **ausente** | — ‖ — |
| Octubre-1 | 2 | 2 `ENEL` | f7 Oportuno#1 = -221,284.53 · f16 Aplicacion#1 = 11,087,940.08 · f27 Oportuno#2 = 13,705,253,661.25 | f42 Aplicacion#2 = -1,847,120.08 · f49 Oportuno#3 = -3,418,096,950.85 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Octubre-1 | 2 | 4 `OCCIDENTE` | f10 Oportuno#1 = -22,430.12 · f22 Aplicacion#1 = 666,142.54 · f37 Oportuno#2 = 791,421,258.90 | f45 Aplicacion#2 = -201,988.54 · f56 Oportuno#3 = 150,601,386.25 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Octubre-1 | 3 | 2 `ENEL` | f7 Oportuno#1 = -558.05 · f13 Oportuno#2 = 9,726,210,483.12 | f28 Oportuno#3 = -1,327,632,993.07 | Main ok / Menor ok ‖ — |
| Octubre-1 | 3 | 4 `OCCIDENTE` | f23 Oportuno#2 = 286,711,907.53 | f35 Oportuno#3 = 145,489,038.47 | — ‖ — |
| Octubre-1 | 4 | 1 `NUEVO ESQUEMA` | f22 Aplicacion#1 = 356,584.67 | f53 Aplicacion#2 = -47,234.67 | — ‖ — |
| Octubre-1 | 4 | 2 `ENEL` | f7 Oportuno#1 = -10,060.29 · f16 Aplicacion#1 = 21,957,966.97 · f35 Oportuno#2 = 10,331,368,139.96 | f50 Aplicacion#2 = -1,560,876.97 · f62 Oportuno#3 = -440,994,165.67 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Octubre-1 | 4 | 4 `OCCIDENTE` | f10 Oportuno#1 = -20,120.65 · f30 Aplicacion#1 = 717,030.70 · f45 Oportuno#2 = 445,264,511.25 | f58 Aplicacion#2 = 133,706.30 · f69 Oportuno#3 = 147,002,644.40 | Main ok / Menor ok / Medio ok ‖ Main ok / Menor ok |
| Octubre-1 | 5 | 2 `ENEL` | f8 Oportuno#1 = 7,414,210,944.89 | f27 Oportuno#2 = -170,561,819.89 | — ‖ — |
| Octubre-1 | 5 | 3 `ENERBIT` | f12 Oportuno#1 = 160,186,886.97 | f30 Oportuno#2 = -20,676,306.97 | — ‖ — |
| Octubre-1 | 5 | 4 `OCCIDENTE` | f22 Oportuno#1 = 134,612,117.14 | f37 Oportuno#2 = 38,861,395.86 | — ‖ — |

**Lectura de la matriz**

- **52 grupos** (= 13 pares empresa×ASE del mapa × 4 períodos) → **180 filas-empresa**. Firma dominante `||<label>|Total|` (A, B y E vacíos; `E` nunca es un concepto no-`Total` → **R3 no se dispara**).
- Cardinalidades observadas en el bloque **Totales**: **1 fila** (14 grupos), **2 filas** (19), **3 filas** (18) y **0 filas** (1: `octubre-1 ASE2 NUEVO ESQUEMA`, ausencia total = Plan 50 A1). En **Subs**: **1 fila** (28), **2 filas** (23) y **0 filas** (1, el mismo grupo ausente). No hay grupos de 4+.
- Todas las secciones son `Oportuno#k` / `Aplicacion#k`; **ninguna fila cae en la cola sin frontera** (Subs/Cont y Total no son filas-empresa) y **ninguna es AFaseo**.
- El **tipo de la primera sección** (la que no tiene frontera arriba) **no es fijo**: cierra con `Aplicacion#1` solo en `julio-1 ASE1` y `julio-2 ASE5`; en las otras 18 fuentes cierra con `Oportuno#1`. La única marca estable de "primera" es *"no tiene frontera arriba"* (`PorFronteraArriba = Primera#0`), nunca el tipo.

---

## 3. S1.2 — Ordinal vs magnitud (fila a fila)

De los **37 grupos con 2+ filas Totales**, **35** asignan los **mismos operandos** por posición y por magnitud (con 2 filas los extremos siempre coinciden: `max |F|` = última sección, `min |F|` = primera). Los grupos de **1 fila** son una clase aparte (§6: la magnitud no puede distinguir `Main` de `Menor` cuando solo hay un bloque).

**Rojo — 2 grupos verdes con magnitud ≠ ordinal (condición de STOP de D3, textual: *"desalineación magnitud≠ordinal en un caso verde actual → STOP + escala"*):**

| Período | ASE | Empresa | Filas Totales (posición) | `TotalMenor` magnitud (hoy) | `TotalMenor` posición (plan) | `TotalMedio` magnitud (hoy) | `TotalMedio` posición (plan) |
|---|---|---|---|---|---|---|---|
| Agosto-2 | 2 | `ENEL` | f7 O1 = −1.414.859,14 · f17 A1 = 131.543,19 · f25 O2 = 24.836.019.402,94 | **131.543,19** (f17) | **−1.414.859,14** (f7) | **−1.414.859,14** (f7) | **131.543,19** (f17) |
| Agosto-2 | 2 | `OCCIDENTE` | f12 O1 = −28.650,84 · f20 A1 = 24.890,00 · f44 O2 = 2.144.338.413,37 | **24.890,00** (f20) | **−28.650,84** (f12) | **−28.650,84** (f12) | **24.890,00** (f20) |

Estos dos casos son **verdes hoy** (el reader devuelve valores sin throw) y en ambos el plan **cambia el operando**. No hay regla de desempate aprobada en el plan: D3 manda STOP.

`TotalMedio` no participa de `CalcularVisibleR1` en ningún ASE (la cadena solo usa `TotalMain`, `EspecialesMain`, `SubsMain`, `TotalMenor` y `Block5Total`), pero `TotalMenor` **sí** ⇒ los dos casos anteriores mueven el visible (§6).

---

## 4. S1.3 — Casos por cardinalidad

**Una fila (14 grupos en Totales):**

| Período | ASE | Empresa | Fila | Sección | ¿Es la primera del bloque? | `Main` (plan) | `Menor` (plan) |
|---|---|---|---|---|---|---|---|
| Julio-1 | 1 | `ENEL` | f13 | Oportuno#1 | no (la 1.ª cierra en `Aplicacion#1`) | valor | 0 |
| Julio-1 | 4 | `NUEVO ESQUEMA` | f34 | Oportuno#2 | no | valor | 0 |
| Julio-1 | 5 | `ENERBIT` | f20 | Oportuno#2 | no | valor | 0 |
| Julio-2 | 2 | `NUEVO ESQUEMA` | f38 | Oportuno#2 | no | valor | 0 |
| Julio-2 | 5 | `ENEL` | f18 | Oportuno#1 | no (la 1.ª cierra en `Aplicacion#1`) | valor | 0 |
| Julio-2 | 5 | `ENERBIT` | f22 | Oportuno#1 | no | valor | 0 |
| Agosto-2 | 2 | `NUEVO ESQUEMA` | f34 | Oportuno#2 | no | valor | 0 |
| Agosto-2 | 5 | `ENEL` | f24 | Oportuno#2 | no | valor | 0 |
| Agosto-2 | 5 | `ENERBIT` | f28 | Oportuno#2 | no | valor | 0 |
| Octubre-1 | 3 | `OCCIDENTE` | f23 | Oportuno#2 | no | valor | 0 |
| Octubre-1 | 4 | `NUEVO ESQUEMA` | f22 | Aplicacion#1 | no (la 1.ª es `Oportuno#1`) | valor | 0 |
| **Octubre-1** | **5** | **`ENEL` / `ENERBIT` / `OCCIDENTE`** | f8 / f12 / f22 | **Oportuno#1** | **sí** | **0** | valor |

- **Octubre-1 ASE3 OCCIDENTE** (el caso del encargo): `f23 = 286.711.907,53`, sección `Oportuno#2`; la hoja tiene **2 secciones**: `Oportuno#1` (solo ENEL, f7 = −558,05) y `Oportuno#2` (ENEL f13 y OCCIDENTE f23) ⇒ **`TotalMain = 286.711.907,53`, `TotalMenor = 0`** — coincide con D3. Hoy el código lanza (`TotalMedio = [1]`), luego este grupo **no es verde** y su Δ queda cubierto por §9-(i).
- **Julio-2 ASE3 OCCIDENTE:** el encargo suponía **1 fila**; el T0 mide **2** (`f7 Oportuno#1 = −16.458,46` y `f28 Oportuno#2 = 825.593.716,53`). Por tanto **D5-test-2 (`NoDuplicado_Jul2_Ase3_Occidente_UnaFilaPrimera`) no aplica** (su condición *"si el T0 confirma 1 fila"* no se cumple) y el Δ esperado es solo `TotalMedio → 0` (§9-ii), no un cambio `Main→0`.
- **Octubre-1 ASE5 (no previsto por el plan):** las 3 empresas traen **una sola** fila y **en la primera sección** ⇒ D3 manda `Main = 0`. La hoja de ASE5-octubre-1 **no tiene sección "primera" de recaudo**: sus 2 únicas secciones son `Oportuno#1` (f8/f12/f22, la única sección de recaudo, que cierra con `Mes#1`) y el bloque de Subs (`Oportuno#2`). El dinero **oportuno** de las 3 empresas está en esa primera sección. Consecuencia: `F483/F497 → 0` y, para `ENERBIT`, **pérdida total** (el mapa `(3,5)` no declara `TotalMenor`): ver §6.

**Dos filas (19 grupos):** `Main` = última sección, `Menor` = primera; `TotalMedio = 0` (hoy `[1]` == `count-1` == `Menor`). Es la clase del §9-ii (`F222` de julio-2/agosto-2 ASE3-OCCIDENTE) y de los duplicados `SubsMenor` de los grupos con **1 sola** sección de Subs (§6).

**Tres filas (18 grupos):** `Main` = última, `Menor` = primera, `Medio` = intermedia por **posición**; coincide con la magnitud en **16 de los 18** — los 2 restantes son exactamente los casos rojos de §3 (`agosto-2 ASE2`).

---

## 5. S1.4 — Inventario R2/R4 (D6, solo conteos)

Conteo de filas-empresa por (período, ASE, empresa del mapa) en las fuentes `RerpoteDetalleSaldosaFavor_*` (R2: `B=label ∧ C=Total`) y `ReversiónPorComponente_*` (R4: `A=label ∧ B=Total`):

- **R2:** 50 de los **52** pares (13 por período × 4) con **exactamente 2** filas (una en el bloque `Componente` y una en `Subs/Cont`). Los **2** pares restantes son **0 filas** (`ENERBIT` ASE5 en julio-1 y en octubre-1) ⇒ ausencia total, ya cubierta por Plan 50 A1 (0 + huella). **Ningún par tiene 1 fila.**
  ⇒ **El defecto gemelo de R2 (`Total = [0]` y `SubsCont = [count-1]` colapsando con 1 fila) NO tiene caso real en los 4 períodos.** Por la decisión ratificada §12-3 ("si el T0 confirma el defecto gemelo en R2 **con caso real** → plan hijo"), **no se requiere plan hijo**; el riesgo queda declarado como latente (una fuente futura con una sola fila-empresa lo dispararía).
- **R4:** todos los pares con 1 fila o 0 filas; **ninguno con 2+** ⇒ `FirstOrDefault` por `A=label ∧ B=Total` no tiene índice de ocurrencia. **Veredicto D6 confirmado: R4 no tiene el defecto.**

---

## 6. S1.5 — Gate de identidad (ROJO)

Definición del gate (D4): para cada (período × ASE × empresa) donde el código actual **devuelve valores sin throw**, la resolución por sección debe reproducirlos **exactos** (`decimal`, no ±0,5). Resultado: **17 operandos-celda divergen** en **13 grupos**; **7 mueven el `VisibleR1`**.

| # | Período | ASE | Empresa | Celda | Operando | Hoy | Plan | Δ | ¿En §9? |
|---|---|---|---|---|---|---|---|---|---|
| 1 | Julio-2 | 1 | `OCCIDENTE` | F29 | SubsMenor | 168.024.399,69 (f46) | 0 | −168.024.399,69 | **NO** |
| 2 | Julio-2 | 3 | `OCCIDENTE` | F242 | SubsMenor | 404.748.758,01 (f44) | 0 | −404.748.758,01 | **NO** |
| 3 | Julio-2 | 3 | `OCCIDENTE` | F222 | TotalMedio | −16.458,46 (f7) | 0 | +16.458,46 | sí (§9-ii) |
| 4 | Julio-2 | 5 | `ENEL` | F474 | TotalMenor | 12.823.196.468,26 (f18) | 0 | −12.823.196.468,26 | **NO** |
| 5 | Agosto-2 | 1 | `OCCIDENTE` | F29 | SubsMenor | 206.290.801,38 (f43) | 0 | −206.290.801,38 | **NO** |
| 6 | Agosto-2 | 2 | `ENEL` | F83 | TotalMenor | 131.543,19 (f17) | −1.414.859,14 (f7) | −1.546.402,33 | **NO** |
| 7 | Agosto-2 | 2 | `OCCIDENTE` | F89 | TotalMenor | 24.890,00 (f20) | −28.650,84 (f12) | −53.540,84 | **NO** |
| 8 | Agosto-2 | 3 | `OCCIDENTE` | F242 | SubsMenor | 360.488.591,23 (f38) | 0 | −360.488.591,23 | **NO** |
| 9 | Agosto-2 | 3 | `OCCIDENTE` | F222 | TotalMedio | −522,80 (f10) | 0 | +522,80 | sí (§9-ii) |
| 10 | Agosto-2 | 5 | `ENEL` | F474 | TotalMenor | 13.260.342.962,39 (f24) | 0 | −13.260.342.962,39 | **NO** |
| 11 | Octubre-1 | 3 | `OCCIDENTE` | F216 | TotalMenor | 286.711.907,53 (f23) | 0 | −286.711.907,53 | sí (§9-i, hoy throw) |
| 12 | Octubre-1 | 3 | `OCCIDENTE` | F242 | SubsMenor | 145.489.038,47 (f35) | 0 | −145.489.038,47 | sí (§9-i, "…+ 0s") |
| 13 | Octubre-1 | 5 | `ENEL` | F483 | TotalMain | 7.414.210.944,89 (f8) | 0 | −7.414.210.944,89 | **NO** |
| 14 | Octubre-1 | 5 | `ENEL` | L483 | EspecialesMain | 10.471.696,87 (f8) | 0 | −10.471.696,87 | **NO** |
| 15 | Octubre-1 | 5 | `ENERBIT` | F487 | TotalMain | 160.186.886,97 (f12) | 0 | −160.186.886,97 | **NO** |
| 16 | Octubre-1 | 5 | `OCCIDENTE` | F497 | TotalMain | 134.612.117,14 (f22) | 0 | −134.612.117,14 | **NO** |
| 17 | Octubre-1 | 5 | `OCCIDENTE` | L497 | EspecialesMain | 1.908.241,19 (f22) | 0 | −1.908.241,19 | **NO** |

**Trece (13) de los 17 Δ no están en la lista cerrada §9** ⇒ por §9 (*"Cualquier otro Δ, alta o baja = STOP"*) el plan queda detenido.

**Impacto en el `VisibleR1`** (`CalcularVisibleR1`, cadena por ASE):

| Período | ASE | Empresa | Visible hoy | Visible con el plan | Δ | Lectura del manual |
|---|---|---|---|---|---|---|
| Julio-2 | 5 | `ENEL` | 24.524.354.313,16 | **11.701.157.844,90** | −12.823.196.468,26 | `REMUNERACION_ENEL!D13 = 11.701.157.844,90` ⇒ **coincide con el plan** |
| Agosto-2 | 5 | `ENEL` | 25.008.738.395,88 | **11.748.395.433,49** | −13.260.342.962,39 | `REMUNERACION_ENEL!D13 = 11.748.395.433,49` ⇒ **coincide con el plan** |
| Agosto-2 | 2 | `ENEL` | 19.239.204.490,16 | **19.237.658.087,83** | −1.546.402,33 | `REMUNERACION_ENEL!D10 = 19.237.658.087,83` ⇒ **coincide con el plan** |
| Agosto-2 | 2 | `OCCIDENTE` | 2.217.333.974,79 | **2.217.280.433,95** | −53.540,84 | `REMUNERACION_OCCIDENTE_ Directa!D10 = 2.217.280.433,95` ⇒ **coincide con el plan** |
| Octubre-1 | 5 | `ENEL` | 14.647.388.373,02 | 7.243.649.125,00 | −7.403.739.248,02 | *sin manual de octubre-1* |
| Octubre-1 | 5 | `ENERBIT` | 139.510.580,00 | **−20.676.306,97** | −160.186.886,97 | *sin manual* — **resultado absurdo (negativo)** |
| Octubre-1 | 5 | `OCCIDENTE` | 306.177.388,95 | 173.473.513,00 | −132.703.875,95 | *sin manual* |
| Octubre-1 | 3 | `OCCIDENTE` | (hoy throw) | 427.389.334,19 | — | *sin manual* |

**Consecuencia de lectura (importante):** los 4 casos con oráculo demuestran que **el defecto es el código actual** (duplica el bloque) y que **la regla por sección es la correcta**. Es decir, D4 no puede exigir identidad "en todo lo verde": en Q2 el verde es falso porque la conciliación por empresa **no se ejecuta** (recorte HU-20/G2-D1) y nadie vio el duplicado. El único bloque donde la identidad sí es exigible hoy es **Q1 con conciliación activa = julio-1** (43 filas-empresa, **0 Δ**) y **octubre-1** (con la excepción de ASE5 y el throw de ASE3).

---

## 7. S1.6 — Tipabilidad, empates y firma (VERDE) + corrección de la premisa D1

- **Tipabilidad:** 180/180 filas-empresa tipadas por `SeccionPorCierre`; **0** devoluciones `null` (nadie es AFaseo), **0** empates (dos filas-empresa de la misma empresa en la misma sección) y **0** secciones intercaladas (una sección que reaparezca después de otra) ⇒ R2-empate y R3 no se disparan.
- **Firma A–E:** todas las filas-empresa son `A=""`, `B=""`, `C=<label>`, `D="Total"`, `E=""`. Ninguna trae concepto distinto de `Total` ⇒ **R3 (E no-`Total`) no se dispara**.
- **Corrección de la premisa D1 (no funcional, sí documental):** D1 afirma que *"para una fila estrictamente interior, cierre y apertura coinciden por construcción"*. **Es falso en 180/180 filas-empresa**: la sección se **abre** por una frontera y se **cierra** por otra (tipos distintos) y hasta en dos `Oportuno` consecutivos los ordinales difieren. Ejemplos:

| Período | ASE | Fila | `SeccionPorCierre` | `PorFronteraArriba` |
|---|---|---|---|---|
| Julio-1 | 1 | f13 (ENEL) | Oportuno#1 | Aplicacion#1 |
| Julio-1 | 1 | f8 (OCCIDENTE) | Aplicacion#1 | Primera#0 |
| Agosto-2 | 5 | f24 (ENEL) | Oportuno#2 | Aplicacion#1 |
| Octubre-1 | 5 | f8 (ENEL) | Oportuno#1 | Primera#0 |

  Lo que **sí** coincide es la **partición** (qué filas caen en la misma sección) y su **orden**: ambas lecturas inducen exactamente los mismos intervalos. Como la resolución de operandos solo necesita (partición, orden, "primera vs no primera"), **el desacuerdo de etiquetas no altera ningún operando**; `SeccionPorCierre` es el tipador vigente (Plan 46 §2.2 = DEFECTO 2, que precisamente existe para que la 1.ª sección no sea un `Primera` genérico). **D1 debe reescribirse citando eso**; no es un STOP funcional.
- **"Primera sección"** (necesaria para D3): la única marca estable es *"sin frontera arriba"* (`PorFronteraArriba = Primera#0`). El **tipo** de la primera sección **varía**: `Aplicacion#1` en julio-1/2, agosto-2 ASE4/5, octubre-1 ASE1/2/4; `Oportuno#1` en agosto-2 ASE3, octubre-1 ASE3/5. Por eso "primera" **no** puede confundirse con "extemporánea/`Aplicacion`": en octubre-1 ASE5 la primera sección **es** el recaudo oportuno (§6).

---

## 8. S1.7 — Arbitraje del oráculo manual

**a) `TotalMenor` duplicado por 1 fila (julio-2/agosto-2 ASE5-ENEL y agosto-2 ASE2):** el manual arbitra **sin ambigüedad a favor de la regla por sección** (4/4):

| Caso | Hoy (ocurrencia) | Plan (sección) | Manual | Veredicto |
|---|---|---|---|---|
| Julio-2 ASE5 `ENEL` | 24.524.354.313,16 | 11.701.157.844,90 | 11.701.157.844,90 | **Plan = manual** |
| Agosto-2 ASE5 `ENEL` | 25.008.738.395,88 | 11.748.395.433,49 | 11.748.395.433,49 | **Plan = manual** |
| Agosto-2 ASE2 `ENEL` | 19.239.204.490,16 | 19.237.658.087,83 | 19.237.658.087,83 | **Plan = manual** |
| Agosto-2 ASE2 `OCCIDENTE` | 2.217.333.974,79 | 2.217.280.433,95 | 2.217.280.433,95 | **Plan = manual** |

Fuente: `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx!REMUNERACION_ENEL!D13` / `Docs/Prueba Agosto-2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx!REMUNERACION_ENEL!D10,D13` y `!REMUNERACION_OCCIDENTE_ Directa!D10`.

**b) `TotalMedio → 0` (julio-2 / agosto-2 ASE3-OCCIDENTE):** **no arbitrable con este oráculo.** `TotalMedio` no entra en `CalcularVisibleR1`, y su único consumidor histórico (`F338 = F242+F222` en la geometría legacy) no existe en la v2: en la plantilla oficial `Reporte Componentes R1!F222`, `!F242` y `!F338` son **slots de valor sin `<f>`** y `CONSOLIDADO!D47:D51` apunta a `F77/F256/F417/F570/F682` (p. ej. `F417 = F342+F315-L315`), cadenas que el mapa por empresa **no** alimenta — de ahí que `CONSOLIDADO!D47:D51 = 0` en las salidas actuales (deuda abierta del Plan 49, no de este plan). El manual (`D47:D51` julio-2 = 49.580 / 56.190.070 / 1.307.680 / 4.343.101 / 441.849) **no** se reproduce con ninguna de las dos hipótesis (`Medio=0` vs `Medio=Menor`) porque su geometría de anclaje es la legacy (R-TRAMPA-ANCLA).
⇒ `Medio = 0` queda **sin contradicción** del manual pero **sin confirmación**: aplica la decisión ratificada §12-1 (divergencia declarada con dueño) **solo si** el Ingeniero acepta el Δ `+16.458,46` (julio-2) y `+522,80` (agosto-2), que son los dos únicos Δ de esta clase.

**c) `SubsMenor` (F29/F242) con 1 sola sección de Subs:** el T0 confirma que en la v2 esas celdas son **hojas inertes** (ningún `<f>` las consume: `F29`, `F242`, `F68`, `F338` no tienen fórmula). Por tanto el Δ es **solo de literal escrito** (sin efecto en CONSOLIDADO ni en `VisibleR1`), pero **sí** altera el diff `values-only` de la salida ⇒ mismo tratamiento que (b): debe estar en la lista cerrada.

---

## 9. Veredicto y condiciones de desbloqueo

### 9.1 Veredicto

**H0 = ROJO.** No se abre H1. Blockchain de la decisión:

| # | Hallazgo | Regla del plan que lo convierte en STOP |
|---|---|---|
| 1 | 13 Δ de operando fuera de la lista cerrada §9 (7 con impacto en `VisibleR1`) | §9 *"Cualquier otro Δ, alta o baja = STOP"* / §12-1 |
| 2 | Magnitud ≠ ordinal en 2 grupos verdes (agosto-2 ASE2 ENEL/OCCIDENTE) | D3 *"desalineación magnitud≠ordinal en un caso verde actual → STOP + escala"* |
| 3 | D3 (1 fila en primera sección ⇒ `Main = 0`) **pierde** 160.186.886,97 en octubre-1 ASE5-ENERBIT (mapa sin `TotalMenor`) y deja `VisibleR1` negativo | No hay decisión ratificada que cubra la pérdida; A2 exige "ningún operando toma el valor de otro", no "el valor desaparece" |
| 4 | Premisa D1 (cierre == apertura) falsa en 180/180 filas-empresa | D1 *"cualquier desacuerdo es STOP"* (se documenta como corrección de justificación, sin impacto funcional) |

### 9.2 Enmiendas propuestas (decisión del Ingeniero, NO auto-aprobadas)

Para que H1 pueda abrirse, el plan necesita **una** de estas dos rutas:

- **Ruta A (recomendada) — ampliar la lista cerrada §9 y acotar D4/D3:**
  1. **§9** pasa a enumerar los 17 Δ de §6 (o, equivalentemente: *"todo `Menor`/`SubsMenor` que hoy duplica al `Main`/`SubsMain` por `count-1` pasa a 0 o al bloque de la primera sección; todo `Main` de un grupo con una sola sección **no primera** se queda"*), con los Δ congelados de la tabla §6 y el arbitraje del manual §8 como evidencia.
  2. **D4** se re-acota: identidad exacta exigible **solo en Q1 con conciliación activa** (julio-1: 0 Δ medido) + los Δ declarados; en Q2 la identidad **no** es exigible porque el verde es falso (conciliación recortada) — con el oráculo del manual como evidencia de que el plan corrige.
  3. **D3 gana una cláusula para la primera sección:** con **una sola** sección presente, el operando es `Main` **si esa sección es la única sección de recaudo de la hoja** (no existe otra sección `Oportuno` distinta) o **si la empresa no declara `TotalMenor` en su mapa**; en caso contrario `Menor`. Alternativa más limpia: **resolver el operando que el mapa declare** (`if (!mapa.Contains(TotalMenor)) ⇒ Main = valor`), que es geometría pura sobre el contrato del mapa (cero ramas por período/empresa, D-H) y elimina la pérdida de octubre-1 ASE5-ENERBIT.
  4. **D1** se reescribe: la coincidencia es de **partición y orden**, no de etiqueta; `SeccionPorCierre` es el tipador único (Plan 46 §2.2).
  5. **§12-1** se extiende a los Δ de `SubsMenor`/`EspecialesMenor` (hojas inertes en v2, §8-c) y a `TotalMedio`.
- **Ruta B — mantener §9/D4/D3 como están:** entonces el plan **no es implementable** sobre estos 4 períodos y debe volver a PROPOSE/DESIGN (la regla por sección seguiría siendo la correcta según el manual, pero el contrato de aceptación sería otro).

**No se ejecutó ninguna de las dos**: H0 es read-only y la enmienda no está aprobada.

---

## 10. Índice de evidencia

| Evidencia | Dónde |
|---|---|
| Matriz completa (52 grupos / 180 filas-empresa) | §2 de este documento |
| Fuentes R1 (20 archivos) | `Docs/Prueba Julio-1/Insumos/{1-Promoambiental,2-Lime,3-Ciudad Limpia,4-Bogota Limpia,5-Área Limpia}/Recaudoporcomponente_*.xlsx` · idem `Prueba Julio-2`, `Prueba Agosto-2`, `Prueba Octubre-1` |
| Filas citadas (todas `Sheet1` de su fuente) | julio-1 ASE1 f8/f13/f23/f28/f32/f39 · julio-1 ASE3 f10/f16/f31/f36/f47 · julio-2 ASE3 f7/f28/f44 · julio-2 ASE5 f18/f43 · agosto-2 ASE2 f7/f17/f25 · agosto-2 ASE3 f10/f26/f38 · agosto-2 ASE5 f24 · octubre-1 ASE3 f23/f35 · octubre-1 ASE5 f8/f12/f22/f27/f30/f37 |
| Oráculo manual julio-2 | `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx!REMUNERACION_ENEL!D13`, `!REMUNERACION_OCCIDENTE_ Directa!D11`, `!CONSOLIDADO_TOTAL RECAUDO!D47:D51` |
| Oráculo manual agosto-2 | `Docs/Prueba Agosto-2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx!REMUNERACION_ENEL!D10:D13`, `!REMUNERACION_OCCIDENTE_ Directa!D10`, `!CONSOLIDADO_TOTAL RECAUDO!D47:D51` |
| Plantilla oficial v2 (slots sin `<f>`) | `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx!Reporte Componentes R1!F29,F68,F77,F222,F242,F338,F417,F568` |
| Inventario R2/R4 | §5 (conteos sobre `RerpoteDetalleSaldosaFavor_*` y `ReversiónPorComponente_*` de los 20 ASE×período) |

**Nada de este documento modifica el repositorio**: es un fixture read-only; el fixture xUnit permanente (`MapeoSeccionOperandosEmpresaR1T0Tests`, plan §7-H1.4) se escribirá cuando la lista §9 esté cerrada, para no congelar como "esperado" un comportamiento que el plan va a cambiar (ni un rojo nuevo en la suite base `p51-h2-suite.trx` = 1260/1193/67/0).
