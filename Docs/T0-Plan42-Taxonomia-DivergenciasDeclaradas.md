# T0 (Plan 42) — Taxonomía de divergencias declarables + evidencia OCCIDENTE ASE5

> **Frente:** Plan 42 (§0/T0, R-T-1, D-A). **Naturaleza:** T0 DIAGNÓSTICO — **cero código de producción**.
> Lectura del validador (`Remuneracion.Core/Services/ValidadorBasico.cs`) + corrida viva read-only del
> path end-to-end desde la plantilla OFICIAL. **El OFICIAL nunca se modifica**
> (SHA256 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).
> **Fixture read-only:** `Remuneracion.IntegrationTests/DivergenciaDeclaradaT0Tests.cs` (congela lo citado
> aquí; no toca producción).

## 0. Propósito

Congelar, leyendo el código real, **qué divergencias puede emitir hoy el validador** (código, celda,
mensaje, gate de origen) y clasificarlas en **declarables** (externas por construcción: discrepancia
oficial-vs-manual con dueño fuera de la app) vs **NUNCA declarables** (defecto interno: lectura,
cálculo, colocación, sellado, naming, preflight, caché). Sin esa tabla cerrada no hay entrada al
catálogo (D-A/R-COLADERA).

## 1. Taxonomía — qué emite hoy el validador

Todos los errores de dominio del validador portan el código `ERR-VALIDACION` (`CodigoError.Validacion`)
y se agregan con el prefijo `[ERR-VALIDACION]`. Los fail-fast de origen (lectura/plantilla) portan sus
propios códigos.

| # | Gate de origen (`ValidadorBasico`) | Código | Celda / sujeto del mensaje | Declarable | Por qué |
|---|---|---|---|---|---|
| 1 | Estructura multi-ASE (consolidados≠leafs, ≠5, ASE duplicado) | `ERR-VALIDACION` | (sin celda; modo) | **NUNCA** | Contrato de entrada interno (composición). |
| 2 | `AjustesSfT` por quincena (Q1=0; Q2=leaf) | `ERR-VALIDACION` | ASE + monto | **NUNCA** | Cálculo/dominio interno. |
| 3 | Gate leaf-vs-consolidado por ASE (R1 F25≈Extemp, R2≈R2TotalOportuno, R4≈ReversionR4) | `ERR-VALIDACION` | ASE + monto | **NUNCA** | Cálculo interno (fuente-fiel). |
| 4 | Σ visibles-empresa == visible de bloque (R1/R2/R4) | `ERR-VALIDACION` | ASE + empresa + monto | **NUNCA** | Cálculo/colocación interno. |
| 5 | DetRetri Q2 (`Detalle == ROUND(D104:D108,0)`) | `ERR-VALIDACION` | ASE + monto | **NUNCA** | Cálculo interno. |
| 6 | GranTotal == Σ TotalAse; TotalAse aritmético | `ERR-VALIDACION` | monto | **NUNCA** | Cálculo interno. |
| 7 | Reporte por banco (C59=quincena; bloque=fuente; consolidado) | `ERR-VALIDACION` | ASE + empresa + monto | **NUNCA** | Lectura/cálculo interno. |
| 8 | Balance SC (TotalBsc=TotalFuente; H≈ROUND(F,0); Σ) | `ERR-VALIDACION` | ASE + monto | **NUNCA** | Lectura/cálculo interno. |
| 9 | Validación por empresa (`VALIDACION_*`: O==0, P==true) | `ERR-VALIDACION` | ASE + empresa | **NUNCA** | Coherencia interna recaudo-vs-remuneración. |
| 10 | `DetValiRetri` (D16:D20==0; D24:D28==true; D29==true) | `ERR-VALIDACION` | ASE + celda | **NUNCA** | Composición interna. |
| 11 | `VALIDACION_TOTAL` O9==0 / P9==true | `ERR-VALIDACION` | ASE + O9/P9 | **NUNCA** | Agregado interno. |
| 12 | **Sub-bloques W-1 `C15`/`D25`** (cadena de recaudo) | `ERR-VALIDACION` | `ASE n · VALIDACION_TOTAL · C15/D25` | **NUNCA** | Cadena 100 % escrita por la app (recaudo). |
| 13 | **Sub-bloques W-1 `O25`/`D34`/`F34`** (cadena del espejo de remuneración) | `ERR-VALIDACION` | `ASE n · VALIDACION_TOTAL · O25/D34/F34` | **Declarable SOLO si el Δ coincide con una entrada del catálogo** | Puede reflejar una discrepancia oficial-vs-manual con dueño externo. Hoy SOLO el par OCCIDENTE ASE5 (Δ = 417.519,62). |
| 14 | DetRetri calculado vs R10 | `ERR-VALIDACION` | ASE + monto | **NUNCA** | Oráculo de período (cálculo interno). |
| 15 | Fail-fast de lectura/colocación/plantilla (`ERR-SIN-SLOT`, `ERR-PLANTILLA`, `ERR-FORMATO-FUENTE`, `ERR-FUENTE-NO-ENCONTRADA`) | propios | fila/hoja/celda | **NUNCA** | Lectura, colocación, naming, preflight, caché — nunca externos. |
| 16 | `ERR-ESCRITURA`, `ERR-INESPERADO`, `WARN-CANCELADO` | propios | (sin celda) | **NUNCA** | I/O, inesperado, decisión del usuario. |

**Regla de cierre (D-A):** lo declarable es una **lista cerrada** = las entradas del catálogo
(`DivergenciasDeclaradasCatalogo`). Todo lo demás es **no declarable por construcción** y sigue el
criterio estricto E1-1..E1-5 (mismo nombre `ASE n · VALIDACION_TOTAL · <celda>`, mismo fail-fast,
misma severidad). Hoy la lista de declarables tiene **una sola entrada**: OCCIDENTE ASE5 (§2).

## 2. Evidencia OCCIDENTE ASE5 — primera entrada del catálogo

### 2.1 Fórmulas (byte-fiel; congeladas por el Plan 41, `T0FrenteEspejoR1Ase5EnelTests`)

| Documento | Celda | Texto-`<f>` | Términos |
|---|---|---|---|
| Oficial | `Reporte Componentes R1!F700` | `F647+F673+F612-L612-L647` | 3 F, 2 L (incluye la 1.ª sección `F612`) |
| Manual legacy | `Reporte Componentes R1!F578` | `F530+F551-L530` | 2 F, 1 L (omite la 1.ª sección) |

El oficial es **internamente consistente** (sus 5 anclas OCCIDENTE tienen 3 términos); el manual
ASE5 es la **anomalía** (sus ASE1–4 sí traen 3 términos). La verdad de negocio la arbitra la UAESP.

### 2.2 Valores (evaluación desde lo escrito; corrida viva Q2)

Colocación por sección: `src13→F612`, `src32→F647`, `src39→F628`, `src53→F673`.

```
Oficial F700 = 229.202.451,62 + 81.691.821,42 + 417.519,62 − 0 − 5.904.652,23 = 305.407.140,43
Manual  F578 = 229.202.451,62 + 81.691.821,42 − 5.904.652,23                  = 304.989.620,81
Δ = 417.519,62 = la fila de la 1.ª sección de OCCIDENTE
```

### 2.3 Divergencia viva en el gate (corrida Q2 2026072 desde el OFICIAL, 5 ASE)

El evaluador W-1 mide los sub-bloques del espejo de remuneración desde lo escrito. Los tres
sub-bloques divergen con **Δ = 417.519,62 exacto** (5 ASE × 3 celdas = 15 veredictos):

| Sub-bloque | Observado | Esperado | Δ |
|---|---|---|---|
| `O25` (`O24=N9`) | 73.139.486.672,75 | 73.139.069.153,13 | 417.519,62 |
| `D34` (`D33=D109`) | 72.441.184.838,33 | 72.440.767.318,71 | 417.519,62 |
| `F34` (`F33=D109+N109`) | 72.584.278.420,00 | 72.583.860.900,38 | 417.519,62 |

`C15`/`D25` (recaudo) cierran TRUE. Los 5 sub-bloques son EVALUABLES desde lo escrito (nunca
"no evaluable" ni verde heredado).

### 2.4 Entrada del catálogo (6 campos citados + Δ)

| Campo | Valor |
|---|---|
| Id | `OCCIDENTE-ASE5` |
| Celda(s) | `O25`, `D34`, `F34` |
| Δ | `417.519,62` |
| Origen del monto | Ancla de OCCIDENTE ASE5 en `Reporte Componentes R1` (oficial 3 términos vs manual legacy 2 términos) |
| Dueño | **UAESP** |
| Cita de evidencia | `Docs/T0-Plan41-MiniT0-OccidenteAse5-SaldosNota-TablaJ.md §1` + `T0FrenteEspejoR1Ase5EnelTests.cs` |
| Fecha | 2026-10-14 |
| Condición de retiro | La UAESP confirme si la remuneración de OCCIDENTE ASE5 incluye la 1.ª sección (se elimina la fila, no se adapta el código) |

## 3. Método y reproducibilidad

- Lectura del código real del validador (`ValidadorBasico`) y corrida viva read-only del path
  end-to-end desde el OFICIAL a una ruta temp FRESCA.
- El fixture `DivergenciaDeclaradaT0Tests` congela: SHA256 del oficial, el Δ exacto por celda, la
  clasificación Declarada con su referencia y el vínculo entrada↔evidencia.
- **Prohibido** `Docs/Insumos` raíz; **prohibido** usar cachés `<v>` como oráculo.
- El oficial se verifica intacto por SHA256 antes/después.
