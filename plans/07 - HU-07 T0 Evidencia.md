# Plan 07 — HU-07: T0 Evidencia y Decisiones (Fase 0)

> Documento de evidencia del Discovery T0 obligatorio (plan §4 Fase 0). Fuente: dump OpenXML raw
> de `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` + lecturas ExcelDataReader de las fuentes
> `REMUNERACION 2026071/{1..5}-*/`. **Nada del mapa multi-ASE entra al código sin esta tabla.**

## 0.1 Clasificación celda-por-celda (valor editable vs fórmula)

Leyenda: `V[val]` = celda VALOR (editable), `F[fórmula]` = celda FÓRMULA (protegida).

### R1 — Reporte Componentes R1

| ASE | TOT_OPT visible | Fórmula | Operandos (todos VALOR) |
|---|---|---|---|
| 1 Promoambiental | F46 `F[F25+F41-L25]` = 16704332434.57 | F25+F41-L25 | F25 V[19549786950.62], F41 V[-2827260776.81], L25 V[18193739.24] |
| 1 | F48 `F[F30+F10-L10]` = 11673020 | F30+F10-L10 | F30 V[5341504.63], F10 V[6331515.37], L10 V[0] |
| 2 Lime | F176 `F[F113+F130-L113+F90-L90]` = 12157780441.19 | F113+F130-L113+F90-L90 | F113 V[16853020231.97], F130 V[-4662382887.59], L113 V[31737181.81], F90 V[-1119721.38], L90 V[0] |
| 2 | **F178 V[0] (sin fórmula)** | — | — (EXTEMP estático 0) |
| 3 Ciudad Limpia | F316 `F[F238+F254+F217-L217-L238]` = 10101988514.82 | F238+F254+F217-L217-L238 | F238 V[11541918400.70], F254 V[-1407588772.89], F217 V[-20882.81], L217 V[0], L238 V[32320230.18] |
| 3 | F318 `F[F243+F223-L223]` = 10198723.07 | F243+F223-L223 | F243 V[4499898.22], F223 V[5698824.85], L223 V[0] |
| 4 Bogotá Limpia | F437 `F[F391+F417+F357-L357-L391]` = 10552409503.83 | F391+F417+F357-L357-L391 | F391 V[11149869380.28], F417 V[-574296773.04], F357 V[-194589.84], L357 V[0], L391 V[22968513.57] |
| 4 | F439 `F[F401+F369-L369]` = 5419780 | F401+F369-L369 | F401 V[130082.13], F369 V[5289697.87], L369 V[0] |
| 5 Área Limpia | F519 `F[F513+F498+F478-L478-L498]` = 8519310329.28 | F513+F498+F478-L478-L498 | F513 V[-184896511.23], F498 V[8731153179.45], F478 V[-12642.22], L478 V[0], L498 V[26933696.72] |
| 5 | **F521 V[0] (sin fórmula)** | — | — (EXTEMP estático 0) |

### R2 — Rem. Anticipos R2 (fórmulas uniformes, operandos VALOR)

| ASE | Visible | Fórmula | Operandos |
|---|---|---|---|
| 1 | E41 = 54216385.68 | E15+E26-K15 | E15 V[56353887.23], E26 V[-2080239.85], K15 V[57261.70] |
| 2 | E135 = 79400801.26 | E85+E103-K85 | E85 V[112733377.55], E103 V[-33332576.29], K85 V[0] |
| 3 | E247 = 31111803.75 | E167+E178-K167 | E167 V[31828083.39], E178 V[-716279.64], K167 V[0] |
| 4 | E343 = 17236000.33 | E290+E308-K290 | E290 V[17671092.07], E308 V[-435091.74], K290 V[0] |
| 5 | E413 = 30774154.23 | E385+E374-K374 | E385 V[-2725453.25], E374 V[33808604.39], K374 V[308996.91] |

### R4 — Reversion Pagos R4 (fórmulas uniformes, operandos VALOR)

| ASE | Visible | Fórmula | Operandos |
|---|---|---|---|
| 1 | D67 = -12054255.65 | D9-P9 | D9 V[-12054255.65], P9 V[0] |
| 2 | D161 = -9889189.72 | D98-P98 | D98 V[-9889189.72], P98 V[0] |
| 3 | D198 = -16103442.89 | D193-P193 | D193 V[-16103442.89], P193 V[0] |
| 4 | D312 = -21288908.57 | D236-P236 | D236 V[-21288908.57], P236 V[0] |
| 5 | D347 = -6421274.68 | D344-P344 | D344 V[-6421274.68], P344 V[0] |

### CONSOLIDADO_TOTAL RECAUDO (todo FÓRMULA, jamás se escribe)

| Celda | Fórmula | Celda | Fórmula |
|---|---|---|---|
| D9 | `'Reporte Componentes R1'!F46` | D28 | `'Rem. Anticipos R2'!E41` |
| D10 | `'Reporte Componentes R1'!F176` | D29 | `'Rem. Anticipos R2'!E135` |
| D11 | `'Reporte Componentes R1'!F316` | D30 | `'Rem. Anticipos R2'!E247` |
| D12 | `'Reporte Componentes R1'!F437` | D31 | `'Rem. Anticipos R2'!E343` |
| D13 | `'Reporte Componentes R1'!F519` | D32 | `'Rem. Anticipos R2'!E413` |
| D47 | `'Reporte Componentes R1'!F48` | D66 | `'Reversion Pagos R4'!D67` |
| D48 | `'Reporte Componentes R1'!F178` | D67 | `'Reversion Pagos R4'!D161` |
| D49 | `'Reporte Componentes R1'!F318` | D68 | `'Reversion Pagos R4'!D198` |
| D50 | `'Reporte Componentes R1'!F439` | D69 | `'Reversion Pagos R4'!D312` |
| D51 | `'Reporte Componentes R1'!F521` | D70 | `'Reversion Pagos R4'!D347` |
| D85..D89 | `'AJUSTES - SF-T'!D47..D51` (valor 0 en Q1) | | |
| D104 | `D9+D28+D47+D66+D85` (maestro shared=47) | D105 | `D10+D29+D48+D67+D86` |
| D106 | shared follower (maestro 47) | D107 | `D12+D31+D50+D69+D88` |
| D108 | `D13+D32+D51+D70+D89` | D109 | `SUM(D104:D108)` |

## 0.2 Decisión F178/F521 (T0-0.2)

- **Hecho**: `R1!F178` (ASE2) y `R1!F521` (ASE5) son celdas **VALOR 0 sin fórmula** en el template golden.
- **Fuentes**: ASE2 y ASE5 SÍ tienen sección extemporánea (`Mes | Total` AFaseo = -1119721.38 / -12642.22),
  pero ese valor alimenta el **operando F90/F478 de la fórmula TOT_OPT**, NO el slot EXTEMP.
- **Golden**: D48 = 0 y D51 = 0 (CONSOLIDADO referencia F178/F521 → 0).
- **Decisión**: el slot EXTEMP de ASE2/ASE5 es un **0 estático del proceso manual**; el writer
  **NO escribe** F178/F521 (no están en el mapa editable por ASE). El 0 del template se conserva
  intacto y el output cuadra con golden D48/D51 = 0. El agregado de dominio `Extemp` HU-02
  (-1119721.38/-12642.22) **no es** el visible EXTEMP; la UI/resumen no debe presentarlo como
  valor CONSOLIDADO (honestidad A5).

## 0.3 Mapeo fuente → template por bloque (T0-0.3)

- **R1**: las fuentes tienen 1..3 filas `Mes|Total` (`colB="Mes"` ∧ `colC="Total"`): ASE1 tiene 2
  (main + Subs/Cont), ASE2..5 tienen 3 (AFaseo + main + Subs/Cont). Labels `Subsidio(-)/Contribucion(+)`
  y `Aplicacion nuevos x reversion` existen en ASE1/ASE3/ASE4; **NO existen en ASE2/ASE5**.
  El reader ASE1-céntrico FALLA en ASE2/ASE5 y produce TOT_OPT incorrecto en ASE3
  (F25+F41-L25 = 11541897517.89 ≠ golden 10101988514.82). → reader por-ASE obligatorio.
- **R2**: labels uniformes (`Componente|Total` → E{a}, `Subs/Cont|Total` → E{b}, col `Especiales` → K{a});
  ASE4 sin columna Especiales → K=0. Solo cambian direcciones de template.
- **R4**: label uniforme (`Total` con B vacío → D{a}, P=0). Solo cambia dirección.

### Mapa editable por ASE (template cell ← slot fuente)

| ASE | R1 TOT_OPT | R1 EXTEMP | R2 | R4 |
|---|---|---|---|---|
| 1 | F25←Mes[0]F, F41←Mes[1]F, L25←Esp(Mes[0]) | F30←Subsidio[0]F, F10←Aplicacion[0]F, L10←0 | E15, E26, K15 | D9, P9 |
| 2 | F113←Mes[1]F, F130←Mes[2]F, L113←Esp(Mes[1]), F90←Mes[0]F, L90←Esp(Mes[0]) | — (estático 0) | E85, E103, K85 | D98, P98 |
| 3 | F238←Mes[1]F, F254←Mes[2]F, F217←Mes[0]F, L217←Esp(Mes[0]), L238←Esp(Mes[1]) | F243←Subsidio[0]F, F223←Aplicacion[0]F, L223←0 | E167, E178, K167 | D193, P193 |
| 4 | F391←Mes[1]F, F417←Mes[2]F, F357←Mes[0]F, L357←Esp(Mes[0]), L391←Esp(Mes[1]) | F401←Aplicacion[1]F, F369←Aplicacion[0]F, L369←0 | E290, E308, K290 | D236, P236 |
| 5 | F513←Mes[2]F, F498←Mes[1]F, F478←Mes[0]F, L478←Esp(Mes[0]), L498←Esp(Mes[1]) | — (estático 0) | E385, E374, K374 | D344, P344 |

### Visibles por bloque (fórmulas T0, aritmética de dominio)

| ASE | TOT_OPT | EXTEMP |
|---|---|---|
| 1 | Mes[0]F + Mes[1]F − Esp(Mes[0]) | Subsidio[0]F + Aplicacion[0]F |
| 2 | Mes[1]F + Mes[2]F − Esp(Mes[1]) + Mes[0]F − Esp(Mes[0]) | 0 |
| 3 | Mes[1]F + Mes[2]F + Mes[0]F − Esp(Mes[0]) − Esp(Mes[1]) | Subsidio[0]F + Aplicacion[0]F |
| 4 | Mes[1]F + Mes[2]F + Mes[0]F − Esp(Mes[0]) − Esp(Mes[1]) | Aplicacion[1]F + Aplicacion[0]F |
| 5 | Mes[1]F + Mes[2]F + Mes[0]F − Esp(Mes[0]) − Esp(Mes[1]) | 0 |

## 0.4 Tabla de evidencia contra golden (verificación aritmética ±0.5)

| ASE | TOT_OPT dominio vs golden | R2 vs golden | EXTEMP vs golden | R4 vs golden | Total fila vs D104:D108 |
|---|---|---|---|---|---|
| 1 | 16704332434.57 ✓ D9 | 54216385.68 ✓ D28 | 11673020 ✓ D47 | -12054255.65 ✓ D66 | 16758167584.60 ✓ D104 |
| 2 | 12157780441.19 ✓ D10 | 79400801.26 ✓ D29 | 0 ✓ D48 | -9889189.72 ✓ D67 | 12227292052.73 ✓ D105 |
| 3 | 10101988514.82 ✓ D11 | 31111803.75 ✓ D30 | 10198723.07 ✓ D49 | -16103442.89 ✓ D68 | 10127195598.75 ✓ D106 |
| 4 | 10552409503.83 ✓ D12 | 17236000.33 ✓ D31 | 5419780 ✓ D50 | -21288908.57 ✓ D69 | 10553776375.59 ✓ D107 |
| 5 | 8519310329.28 ✓ D13 | 30774154.23 ✓ D32 | 0 ✓ D51 | -6421274.68 ✓ D70 | 8543663208.83 ✓ D108 |
| GranTotal | — | — | — | — | 58210094820.50 ✓ D109 |

## 0.5 Confirmaciones (T0-0.5)

- D85:D89 = fórmulas → `'AJUSTES - SF-T'!D47..D51` con valor 0 en Q1 ✓ (no se toca; 2.5 es otra HU).
- D104:D108/D109 fórmulas protegidas; D106 es shared follower del maestro D104 (shared=47) —
  la validación debe tolerar followers con texto vacío + SharedIndex.
- ASE4 sin columna Especiales en fuente R2 (TieneColumnaEspeciales=False) → K290=0 ✓ re-verificado.

## 0.6 Consecuencias de implementación

1. `WorkbookLeafCellMapPorAse` congelado según §0.3 (editables + visibles protegidos por `Ase.Id`).
2. `ExcelDataReaderWorkbookLeafInputReader` se extiende a por-ASE (contrato intacto; implementación
   nueva por bloque). Se conservan las propiedades ASE1 para compatibilidad single-ASE.
3. `ValidarContraFuentes` sigue aplicando (F25=Mes[0]F=Extemporaneo HU-02 en todos los ASE).
4. Fórmulas CONSOLIDADO D9:D13/D28:D32/D47:D51/D66:D70/D85:D89/D104:D109 ampliadas en validación
   protegida (shared-follower awareness para D106).

---

## 5.1 Capa B manual — evidencia (protocolo §5.3 del plan, ejecutado el 2026-09-08)

Salida generada con el procesador real (`ProcesadorPeriodo`, insumos Q1) a ruta temporal distinta
del golden. Se leyeron los 26 valores CONSOLIDADO de la SALIDA (caché OpenXML, no recálculo Excel)
y se compararon contra el golden ±0.5:

| Grupo | Celdas | Resultado |
|---|---|---|
| TOT_OPT | D9..D13 | 5/5 OK, diff 0.00 |
| R2 | D28..D32 | 5/5 OK, diff 0.00 |
| EXTEMP | D47..D51 | 5/5 OK, diff 0.00 (ASE2/ASE5 = 0) |
| R4 | D66..D70 | 5/5 OK, diff 0.00 |
| Totales | D104..D108 | 5/5 OK, diff 0.00 |
| GranTotal | D109 | 58210094820.50 OK, diff 0.00 |

**Nota honesta:** la comparación de valores de la salida usa la caché OpenXML de la SALIDA; no
sustituye el recálculo real de Excel (Capa B residual). No es gate de merge (plan §5.3). GranTotal
post-Excel = Σ visibles leaf por ASE = 58210094820.50 = D109 golden.

## 5.2 Cierre de alcance (plan §5.2)

FUERA de HU-07, para HUs posteriores (§9 Fase 2 it. 2.2–2.7):
- 2.2 conciliación por empresa (hojas EAAB/ENEL/ENERBIT/OCCIDENTE/EAAB-CL, VALIDACION_*).
- 2.3 REPORTE RECAUDO x BANCO. 2.4 BCE SC POR FACT.
- 2.5 AJUSTES-SF-T (SALDOS POR NOTA + RETRIBUCION NEGATIVA; Q1 sigue AjustesSfT=0; quincena 2 sigue
  bloqueada en `CalculoRemuneracion`).
- 2.6 DetRetri / DetValiRetri (enteros como objetivo de escritura).
- 2.7 validaciones cruzadas (VALIDACION_TOTAL, VALIDACION_RECIP, etc.).

Próxima HU propuesta: **2.2 conciliación por empresa**.