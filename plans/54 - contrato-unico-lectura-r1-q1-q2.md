---
titulo: Plan 54 - Contrato único de lectura R1 Q1/Q2 (fin de la doble resolución por quincena)
estado: CERRADO 2026-10-22 (implementado H0→H3; code-review independiente APROBADO con observaciones, 0 bloqueantes)
fecha: 2026-10-22
numero: 54
planes_ocupados: 01..53 ocupados; este es el 54
padre: plans/53 - contrato-seccion-por-operando-r1.md
madre_doctrina: plans/50 §0.2 punto 4 ("Q1 y Q2 se tratan igual: no hay reglas particulares por quincena") + `.opencode/project-context.md` (regla 3: la quincena NO selecciona algoritmo; solo fuentes y declaración de visibles)
caso_rojo: corrida real 2026101 (octubre-1, Q1), plantilla v2 — `CalculoInvalidoException — ASE 3: no se encontró la fila 'Aplicacion nuevos x reversion' del R1 para F223.`
arbitraje_r-oraculo: RESUELTO 2026-10-22 por el Ingeniero con bytes v2 (EXTEMP ASE3/ASE5 = 0,00; valor −1.327.632.993,07 descartado como semántica v1; ver §0.2-F12/F13, §1.4 R-ORACULO, §3-CA, §6)
enmienda_1: 2026-10-22 — CA2/CA3/D6 re-anclados al R10 oficial `R10_Remuneracion_2026101.xlsx` tras STOP del implementer en T0.3: ASE1 14.842.425.136 / ASE2 11.302.637.605 / ASE3 TOT_OPT 8.811.472.160,34 / ASE5 TOT_OPT 7.544.253.279,94; ASE5 deja de ser cuarentena (R10 lo valida con 2 filas). Se añade CA0 (todo criterio con derivación) y H1.7/S2.6 (hueco de guard `F682`). Evidencia: `Docs/T0-Plan54-Contrato-Unico-Lectura-R1.md`
enmienda_2: 2026-10-22 — ratificada tras Δ-1 del T0 (julio-1 ASE1 con 2 filas `Mes/Total` y `Mes2` declarado core): `Mes2`/`Lmes2` pasan a opcionales → `0m`; `Mes0`/`Mes1` siguen core con mensaje byte-idéntico. Cero cambio de valor: julio-1 reproduce su R10 al peso (16.758.167.585). Ver D1.1
gate_entrada:
  plantilla_v2: NO SE TOCA (SHA `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` — medido 2026-10-22; verificar intacto antes/después)
  plantilla_v1: NO SE TOCA (histórica; SHA `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` intacto antes/después)
  alcance_aprobado: T1 (resolución R1 única) + T2 (mapa R1 único) + T3 (R2/R4 tabla única) + T4 (candado anti-recaída); prohibido ampliar sin nuevo OK del Ingeniero
  suite_base: p53-h2-suite.trx (1289 Total / 1222 Passed / 67 Failed / 0 Skipped — Plan 53 §13)
  sin_commits: NUNCA `git commit` / `git push`
---

# Plan 54 — Contrato único de lectura R1 Q1/Q2

> SDD completo PROPOSE → DESIGN → SPEC → TASKS. Fuente única de verdad para la implementación una vez aprobado ESTE documento. No se implementa sin aprobación del Ingeniero (punto de bloqueo: sin su OK no se abre H0).
> Lectura obligatoria previa: `plans/50` (§0.2 doctrina, §2.4, §8 residuo R-MAPA-V2-R2-ASE4) + `plans/25` (ruta por firma, agregados totOpt/extemp) + `plans/41` (E1-1/E1-3/E1-6, ERR-SIN-SLOT) + `plans/45`/`plans/46` (ContratoSeccionPeriodFreeR1, SeccionPorCierre) + `plans/48` (adopción v2, oráculo julio-2) + `plans/49` (gate absoluto, celda-referenciada) + `plans/53` (contrato por operando, lista §9, desviaciones ratificadas) + `.opencode/project-context.md` (doctrina transversal + Planes 21-53).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Contexto adoptado (confirmado en código el 2026-10-22; no re-investigar)

### 0.1 El rojo que motiva el plan

Corrida `2026101` (octubre-1, Q1) aborta en ASE 3 (Ciudad Limpia):

```text
CalculoInvalidoException — ASE 3: no se encontró la fila 'Aplicacion nuevos x reversion' del R1 para F223.
```

Es el `?? throw` de `MapearR1` (`ExcelDataReaderWorkbookLeafInputReader.cs:1503`, `case 3`): la rama Q1 exige filas `Aplicacion nuevos x reversion` con ordinales congelados del golden v1 de julio-1, y la fuente octubre-1 de ASE3 trae **cero** filas `Aplicacion` (fuente internamente consistente — su fila `Total` cierra al centavo, §0.2-F4). La rama Q2 (`MapearR1Q2`) ya absorbe ese caso por diseño (Aplic opcionales → 0). El defecto no es la fuente: es la **doble resolución**.

### 0.2 Hechos verificados (archivo:línea, confirmados el 2026-10-22)

| # | Hecho | Cita |
|---|---|---|
| F1 | Dispatch por quincena en el reader leaf: `var esQuincena2 = periodo.NumeroQuincena == 2;` y `R1 = esQuincena2 ? MapearR1Q2(...) : MapearR1(...)` (ídem R2/R4, líneas 50-51) | `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs:43-52` |
| F2 | `MapearR1` (1413+) resuelve con `switch (ase.Id)` (1447) + ordinales congelados: `case 1` exige `filasAplicacion.FirstOrDefault() ?? throw` para F10 (1459-1460); `case 3` exige fila Aplicacion para F223 (1503); `case 4` exige 2 filas Aplicacion para F401/F369 (1534); rellena `F30/F10` vía helpers (1575-1576) | `ExcelDataReaderWorkbookLeafInputReader.cs:1413-1600` aprox. |
| F3 | `ObtenerF30` (2042-2059) y `ObtenerF10` (2061+) ramifican por ASE; uso exclusivo dentro de `MapearR1` (1575-1576) — verificar por grep en H1, sin otro consumidor esperado | `ExcelDataReaderWorkbookLeafInputReader.cs:2042-2080` aprox. |
| F4 | `MapearR1Q2` (1214-1312) YA es la forma correcta: secuencia tipada (`ConstruirFilaEspejoRolQ2` 1224-1228), `FilasMes()`/`FilasAplicacion()` (1230-1231), `RolCore` Mes0..2/Lmes0..2 (1236-1245), Aplic opcionales → 0 (1247-1260, 1273-1275), fail-fast SOLO core (1252-1260, mensaje intacto); `totOpt = Σ Mes − Esp(todas menos la última)` (1286-1289); `extemp = Σ Aplic − Esp(Aplic[0])` (1292-1293); `F25/F41/L25` primera/segunda fila Mes (1295-1304); `F30/F10/L10 = 0m` constantes (1305-1307) | `ExcelDataReaderWorkbookLeafInputReader.cs:1214-1312` |
| F5 | Insumos octubre-1 (censo del encargo, el T0 lo congela por bytes): ASE3 `3-Ciudad Limpia/Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174753997.xlsx`, hoja `Sheet1`: filas `Mes/Total` 8/24/36, CERO `Aplicacion`, fila 38 `A=Total` = 8.830.777.878 = 10.012.922.390,65 − 1.182.143.954,60 − 558,05. ASE5 `5-Área Limpia/..._20261016175148433.xlsx`: solo 2 `Mes/Total` (23 y 38), fila 40 `A=Total` = 7.556.633.218 = 7.709.009.949 − 152.376.731. ASE1 (18, 41), ASE2 (23, 46), ASE4 (31, 59) sí traen 2 `Aplicacion` | `Docs/Prueba Octubre-1/Insumos/...` (rutas del encargo) |
| F6 | Plantilla v2 `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx`, hoja `Reporte Componentes R1` (sheet10.xml). Fronteras por bloque: ASE1 16/32/47/59/70 · ASE2 120/146/171/190/208 · ASE3 299/315/330/342/353 · ASE4 460/486/511/530/548 · ASE5 613/629/648/660/674. TOT_OPT/EXTEMP leídos de los bytes (resolución 2026-10-22): ASE1 `F75=<f>F47+F70+F16-L16-L47</f>` · `F77=<f>F59+F32-L32</f>`; ASE2 `F254=<f>F171+F208-L171+F120-L120</f>` · `F256=<f>F190+F146-L146</f>`; ASE3 `F415=<f>F330+F353+F299-L299-L330</f>` · `F417=<f>F342+F315-L315</f>`; ASE4 `F568=<f>F511+F548+F460-L460-L511</f>` · `F570=<f>F530+F486-L486</f>`; ASE5 `F680=<f>F674+F648+F613-L613-L648</f>` · `F682=<f>F660+F629-L629</f>` ⇒ **v2 homogénea (3 Mes + 2 Aplic por ASE) y el EXTEMP es SIEMPRE `Σ Aplic − Esp(Aplic[0])`, sin término de subsidio**; el "literal 0 sin fórmula" (T0-0.2/Plan 21, V0.2/Plan 25) era verdad de la v1 y YA NO aplica en la v2 vigente (`0F192AD2…`, reemplazo `d55e621` documentado en Plan 50 §8). Coincide exactamente con `extemp` de `MapearR1Q2` (1292-1293) | v2 por bytes (el T0.1 adopta esta tabla o la corrige con bytes) |
| F7 | ESCRITURA R1 ya unificada y NO SE TOCA: `OpenXmlPlantillaWriter.GenerarWorkbook` (246-302) usa `EscritorValoresSlotR1` (contrato por sección, values-only `<v>`, sin insert/delete, sin `<f>`, guard `ERR-PLANTILLA`) y luego `EscribirCeldasLeafPorAse(..., omitirR1: espejoR1)` (312-313; skip del bloque R1 en 1064-1108) | `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs:246-313,1052-1108` |
| F8 | Mapas: `WorkbookLeafCellMapQ2.R1Q2EditablesPorAse` (71+) es la geometría real de la v2; `WorkbookLeafCellMapPorAse.EditableLeafCellsPorAse` (23-63) es geometría v1 (F25/F41/L25/F30/F10/L10) | `WorkbookLeafCellMapQ2.cs:71+`, `WorkbookLeafCellMapPorAse.cs:23-63` |
| F9 | `MapearR2` (1958-2012) y `MapearR4` (2014-2040) YA son label-based en lectura (filas `Componente/Total` 1962-1965, `Subs/Cont/Total` 1967-1970, `Total` R4 2016-2019; columna E por índice 4, Especiales por encabezado); su `switch (ase.Id)` (1978-2003) solo elige **celdas destino** (`E21/E32`, `E101/E123`, `E193/E204`, `E327/E349`, `E426/E440` — coinciden con la v2 medida, hecho 5 del encargo). Gemelas Q2: `MapearR2Q2` (1346-1381, vía `ObtenerR2Q2Editables` 1366), `MapearR4Q2` (1388+) | `ExcelDataReaderWorkbookLeafInputReader.cs:1346-1381,1958-2040` |
| F10 | Consumidores reales de `F25/F41/L25`: `WorkbookLeafCoherence` + `ValidadorBasico.cs:812` (`F25` vs `Extemp`). Consumidores de `F30/F10/L10`: solo path legacy `EscribirDetalleR1` (`OpenXmlPlantillaWriter.cs:1014-1019`), `Herramientas/VerificadorRecaudo/Program.cs:360-367,401`, `ParidadCliTests`, `WorkbookLeafInputsTests` (inventario inicial por grep; el T0.5 lo cierra exacto) | citas en cada archivo |
| F11 | Recorte Q2 vigente (residuo, NO se toca): `Remuneracion.Core/Services/ProcesadorPeriodo.cs:198-217` (`if (!esQuincena2)` omite conciliación por empresa en Q2; deuda `R-MAPA-V2-R2-ASE4`: `Rem. Anticipos R2!E285`/`K285` fórmula en v2, `E285=E192+E203-K192`) | `ProcesadorPeriodo.cs:198-217` |
| F12 | `MapearR1` calcula el EXTEMP con **semántica v1** (`extemp = F30+F10−L10` = subsidio[0] + Aplic[0] − 0 en ASE1/ASE3; `Aplic[1]+Aplic[0]` en ASE4), que **NO** coincide con la fórmula v2 `Σ Aplic − L(Aplic[0])` (F6). Es una divergencia adicional del defecto y refuerza D1: ni siquiera donde la rama Q1 no lanza, calcula lo que la v2 declara. Ejemplo ASE1-octubre (el T0.3 lo deriva al centavo): subsidio[0] fila 35 = −32.099,64; Aplic fila 18 = 636.047,07 y fila 41 = 12.822,93 | código 1447-1600 + fuente ASE1 octubre-1 |
| F13 | Arbitraje R-ORACULO RESUELTO por bytes v2 (2026-10-22): el EXTEMP v2 es siempre `Σ Aplicacion − Especiales(Aplic[0])`. ASE3-octubre trae CERO filas `Aplicacion` ⇒ slots `F315`/`F342`/`L315` en el 0 del template ⇒ `F417 = 0+0−0 = 0,00`; ídem ASE5 (`F682 = 0,00`). El valor **−1.327.632.993,07** queda descartado: era semántica v1 (`F318 = F243+F223−L223`, con término de subsidio) aportada por el arquitecto y refutada por los bytes; en este documento solo se cita como "valor v1 descartado por evidencia" | v2 sheet10.xml por bytes + §6 |

### 0.3 Doctrina madre (no se re-litiga)

`plans/50` §0.2 punto 4 + `.opencode/project-context.md` regla 3: **Q1 y Q2 se tratan igual; no hay reglas particulares por quincena. La quincena NO selecciona algoritmo; solo gobierna fuentes y declaración de visibles.** Usos legítimos supervivientes de `NumeroQuincena` (allow-list cerrada, §2-D7): detalle R2/R4 solo Q2 (Plan 29), `AjustesSfT = 0` en Q1 (regla de negocio), `RESUMEN MES` D/E-vs-F/G (dominio), `L9:L13` + F-NEW-1 por variante (Planes 39/40), DetRetri Q2-escribe/Q1-valida (G3), naming por período (Plan 30), preflight saldos/retribución Q2-only (Planes 26/27). Todo lo demás que ramifique por quincena es el defecto que este plan elimina.

---

## 1. PROPOSE

### 1.1 Intent

Que la lectura R1 tenga **una sola resolución por firma/sección para ambas quincenas**: la forma `MapearR1Q2` (firmas, roles core-fail-fast + Aplic opcionales, agregados Σ) absorbe a `MapearR1` (switch por ASE + ordinales congelados v1), que se elimina. El mapa productivo R1 pasa a ser único (geometría v2). R2/R4 colapsan a tabla única (ya son label-based en ambas ramas). Un candado (test) impide que el lector leaf vuelva a ramificar por `Periodo.NumeroQuincena`. La corrida `2026101` corre los 5 ASE sin excepción; `2026071` y `2026072` salen verdes **por la misma ruta**; la plantilla v2 queda byte-idéntica salvo los 11 tokens de período aprobados.

### 1.2 In Scope (T1–T4, prohibido ampliar sin nuevo OK del Ingeniero)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **T1** | Una sola resolución R1: eliminar `MapearR1` y su `switch (ase.Id)`; la resolución por firma/sección queda como única para ambas quincenas. Desaparece el dispatch por `NumeroQuincena` en R1 (línea 49). Destino explícito de `F30/F10/L10` (legacy) sin reintroducir rama de quincena (§2-D2) | `2026101` ASE3 deja de lanzar; Q1 y Q2 resuelven por el mismo método; grep de cierre sin `MapearR1`-viejo ni `ObtenerF30/ObtenerF10` |
| **T2** | Mapa R1 único: `R1Q2EditablesPorAse` (geometría v2) es el mapa para ambas quincenas; `EditableLeafCellsPorAse` se retira del path productivo y se aísla para el API legacy (§2-D3) | Ningún path productivo referencia el mapa v1; writer sin rama-R1 por quincena |
| **T3** | R2/R4 a tabla única: colapsar `MapearR2/MapearR2Q2` y `MapearR4/MapearR4Q2` a un solo método + una sola tabla destino por reporte, **sin cambiar valores** (precondición T0.4: igualdad celda-por-celda probada antes de borrar) | Salidas Q1/Q2 R2/R4 byte-idénticas antes/después del colapso (caracterización) |
| **T4** | Candado anti-recaída: test que falla si el lector leaf vuelve a ramificar por `Periodo.NumeroQuincena` (estático allow-list + conductual Q1≡Q2, §2-D5) | Test verde que enrojece ante cualquier `NumeroQuincena` nuevo en el reader fuera de allow-list |

### 1.3 Non-goals / out (vinculante)

- No tocar `EscritorValoresSlotR1`, `ContratoSeccionPeriodFreeR1` (`SeccionPorCierre`/`Alinear`), `PreparadorBasePipeline`, ni ninguna `<f>`/geometría/nombre de hoja de la plantilla (11 tokens de período sobre la COPIA, ya aprobados).
- No tocar el recorte `ProcesadorPeriodo.cs:198-217` (residuo §7, con su evidencia y su gate de reactivación).
- No corregir EXTEMP/R2 ni reclasificar divergencias ajenas (Planes 49/50/53 y sus deudas intactas).
- No reabrir la semántica de agregados (`totOpt`/`extemp` Plan 25 D-C/D-F intactos) ni el contrato por operando de conciliación (Plan 53 intacto).
- Commits, UI, CLI, docs de usuario.

### 1.4 Riesgos con mitigación

| ID | Riesgo | Mitigación |
|---|---|---|
| R-Q1-SILENCIO | Al borrar la rama Q1, un comportamiento Q1 real (no solo el throw) cambia sin que nadie lo note | T0.6 congela la lista cerrada de Δ; caracterización Q1-vs-Q2 antes del borrado (mismo fixture, ambas quincenas, diff citado); rojo fuera de lista = STOP (§3-S0) |
| R-F30 | `F30/F10/L10` tienen consumidores vivos (legacy + herramientas + tests) y su borrado ingenuo rompe compilación o paridad | D2: miembros intactos, valor constante 0 (semántica Q2 actual), legacy aislado con marca explícita, re-anclaje con evidencia (§2-D7) |
| R-TABLA | `MapearR2Q2`/`MapearR4Q2` difieren de sus gemelas Q1 en algo más que la tabla destino (valor, columna, tolerancia) | T0.4 bloqueante: caracterización con fixtures reales julio-1+julio-2; cualquier diferencia de VALOR = STOP, nunca remap silencioso |
| R-RECAIDA | Un fix futuro reintroduce `if (quincena)` en el reader por comodidad | T4: candado de dos capas (grep allow-list + identidad conductual Q1≡Q2) |
| R-ORACULO — RESUELTO 2026-10-22 | El criterio CA3 (−1.327.632.993,07) era semántica v1 aportada por el arquitecto; refutado por bytes v2 (§0.2-F13): EXTEMP v2 = `Σ Aplic − Esp(Aplic[0])` ⇒ ASE3/ASE5 = 0,00. Registro de la corrección: criterio aportado → bytes del template → descarte documentado (el valor solo sobrevive citado como "v1 descartado") | T0.3 verifica contra los criterios corregidos (§3-CA); mismatch ⇒ STOP, nunca tolerancia |
| R-ASE5 | El slot `Mes2`/`F674` en 0 se lee como "verde" y una 3.ª fila futura se absorbe en silencio | D6: divergencia declarada con 6 campos + test que la pincha; agregados Σ intactos absorben filas extra por regla documentada (Plan 25), nunca por slot (ver §2-D6) |
| R-PLANTILLA | El colapso de mapas/818 escrituras toca una celda-fórmula | Guard `ERR-PLANTILLA` intacto + SHA v2 antes/después (CA6) + `<f>` byte-idéntica en el diff (G-FORMULA) |

---

## 2. DESIGN

### D1. Ruta única R1: `MapearR1Q2` absorbe a `MapearR1` (decisión)

**Forma normativa:** se elimina `MapearR1` (`ExcelDataReaderWorkbookLeafInputReader.cs:1413` + `switch` 1447 + ordinales) y se renombra `MapearR1Q2` → `MapearR1` (mismo cuerpo 1214-1312, ajustado solo en nombre/mensajes — ver D1.3). La línea 49 pasa a `R1 = MapearR1(ase, filasR1)` sin `esQuincena2`. `ObtenerF30`/`ObtenerF10` (2042+) se eliminan (muertos tras el switch; H1 lo prueba por grep).

| Candidata | Veredicto | Motivo |
|---|---|---|
| **U-A: renombrar Q2→único, borrar Q1** | **Elegida** | El cuerpo correcto ya existe, probado (Planes 25/53, goldens, R10 10/10). Diff mínimo y direccional: solo borra lo viejo. El nombre `MapearR1Q2` mentiría tras la unificación |
| U-B: tercer método nuevo + borrar ambos | Descartada | Duplica una implementación probada para renombrarla; riesgo de divergencia accidental sin ganancia |
| U-C: mantener ambos, Q1 delegando en Q2 | Descartada | Deja el switch muerto como "documentación" y conserva la rama que el candado T4 debe prohibir; confunde al próximo lector |

**D1.1 Semántica única (enmendada 2026-10-22 — enmienda 2):** secuencia tipada → `Mes0`/`Mes1` core con fail-fast (mensaje de las líneas 1257-1258 byte-idéntico) → **`Mes2`/`Lmes2` opcionales → `0m`** (mismo trato que `Aplic0/Aplic1/LAplic0`) → `totOpt` Σ (1286-1289) → `extemp` Σ (1292-1293) → `F25/F41/L25` primera/segunda Mes (1302-1304). Cero cambios de aritmética.

**Motivo de la enmienda 2 (Δ-1 del T0):** julio-1 ASE1 trae 2 filas `Mes/Total` y el mapa v2 le declara `Mes2 → F70`; con `Mes2` core la ruta única abortaba y CA4 era inalcanzable. Con `Mes2 → 0` la ruta reproduce el R10 oficial de julio-1 **al peso** (`16.758.167.585`) y el `TotalOportunoEsperado` v1 de esa fuente coincide exacto ⇒ **cero cambio de valor**. Generaliza la doctrina ya ratificada: **el slot que la fuente no llena queda en el cero del template**. No se hacen opcionales `Mes0`/`Mes1` (se conserva la guarda de forma mínima); la clase completa "slot declarado sin fila fuente" con su control anti-falso-0 queda para plan propio (Plan 55).

**D1.2 Mensajes de error:** el texto fail-fast core se conserva byte-idéntico (hay tests que lo asertan). Los `throw` del `switch` eliminado desaparecen con él; cada uno queda trazado en el T0.6 (qué aborto Q1 deja de existir y por qué la ausencia correspondiente hoy es 0 documentado o fail-fast core).

**D1.3 Nota de nombres:** `R1Q2Fuente` (enum del mapa) conserva su nombre en este plan (renombrarlo churnea mapa + tests sin ganancia conductual); se documenta en su doc-comment que rige ambas quincenas. Renombrar el enum queda como follow-up explícito fuera de alcance.

### D2. Destino de `F30/F10/L10` (legacy) — decisión explícita sin rama de quincena

**Forma normativa:** los tres miembros permanecen en `WorkbookLeafInputsR1` (compatibilidad de contrato) y la ruta única los puebla con la constante `0m` en ambas quincenas — exactamente la semántica Q2 actual (1305-1307). Constante, nunca rama.

| Candidata | Veredicto | Motivo |
|---|---|---|
| **L-A: constante 0 + miembros intactos + legacy aislado** | **Elegida** | Cero cambio conductual Q2; Q1 pasa de "valor ordinal-congelado" a 0 con trazabilidad T0.6. Ningún path productivo los lee (F7 + F10) |
| L-B: derivarlos de roles (`F10=Aplic0`, …) | Descartada | Cambiaría Q2 (hoy son 0) y les inventaría una semántica que ningún consumidor productivo usa; viola T3-"sin cambiar valores" por analogía |
| L-C: eliminar los miembros | Descartada | Rompe `VerificadorRecaudo`, tests y el path `EscribirDetalleR1`; la eliminación es un plan propio con su T0 |

**Aislamiento legacy:** el path `EscribirDetalleR1` (`OpenXmlPlantillaWriter.cs:1014-1019`) se marca `LEGACY-only, fuera de paridad` (doc-comment + `[Obsolete]` si el target lo permite sin warnings — si genera ruido, solo doc-comment; 0 warnings es inviolable). Prohibido añadir lectores nuevos de `F30/F10/L10` (grep de cierre). `Herramientas/VerificadorRecaudo` sigue compilando intacto (herramienta de diagnóstico, no camino del writer — doctrina Plan 28).

### D3. Mapa R1 único (decisión)

**Forma normativa:** `R1Q2EditablesPorAse` es el mapa productivo para ambas quincenas. `EditableLeafCellsPorAse` + su helper (`WorkbookLeafCellMapPorAse.cs:141`) se retiran del path productivo y se aíslan con marca `LEGACY (geometría v1)` para el API legacy. Consumidores productivos a re-apuntar (inventario exacto en T0.5): el branch-Q1 del writer (`EscribirCeldasLeafPorAse`, 1052-1062) y el iterador de la línea 759 (el T0.5 cita en qué método vive y su veredicto). Tras el re-apunte, la rama-R1 por quincena del writer colapsa (si solo servía a R1, se elimina; si comparte branch con R2/R4, la porción R1 sale del `if` con cita).

### D4. R2/R4 a tabla única (decisión condicionada)

**Forma normativa (si T0.4 prueba igualdad):** un solo `MapearR2` + una sola tabla destino (el `switch` 1978-2003, cuyos valores coinciden con la v2 medida), un solo `MapearR4` + su tabla; se eliminan `MapearR2Q2`/`MapearR4Q2` (+ `ObtenerR2Q2Editables` y gemela si el T0.4 las declara idénticas). Líneas 50-51 colapsan a llamada directa. **Condición vinculante:** la caracterización T0.4 debe probar igualdad de VALORES (no solo de celdas) Q1-vs-Q2 sobre fixtures reales julio-1 + julio-2 (5 ASE × 2 quincenas). Cualquier Δ de valor = STOP (R-TABLA), nunca "armonización".

**Por qué es seguro por construcción:** ambas ramas leen por label las mismas filas de cierre (F9: mismos predicados, mismos índices de columna 4/Especiales-por-encabezado); la única divergencia declarada es la tabla destino, y el encargo ya midió que la v2 trae esas celdas en ambas variantes. El T0.4 lo convierte en prueba, no en supuesto. Las celdas destino deben probarse literales escribibles por bytes (sin `<f>`; si alguna trae fórmula → STOP, clase `R-MAPA-V2-R2-ASE4`).

### D5. Candado anti-recaída T4 (decisión)

Dos capas, un test (`LecturaR1SinRamaQuincenaTests` o nombre que fije H1):

1. **Estática (allow-list):** el cuerpo de `ExcelDataReaderWorkbookLeafInputReader.cs` no contiene `NumeroQuincena` salvo allow-list cerrada: construcción `DetalleR2/DetalleR4` solo-Q2 (58-62, alcance Plan 29 — fuera de este plan) y ninguna otra tras T1–T3. Implementación: test que lee el fuente como texto y aserta `Regex.Matches(...) == allow-list` con cita de línea (patrón: gate de invariante estructural, no de negocio; si el archivo se reordena, el test se re-ancla, no se borra).
2. **Conductual (identidad):** el mismo fixture R1 (octubre-1 ASE3 + julio-1 ASE1) resuelto con `Periodo` Q1 y Q2 produce `WorkbookLeafInputsR1` idénticos miembro a miembro (`F25/F41/L25`, `CeldasPorAse`, ambos agregados). Es la prueba de que "la quincena no selecciona algoritmo".

### D6. ASE5-octubre: forma de 2 filas `Mes/Total` — VALIDADA por el R10 (corrección 2026-10-22)

El R10 oficial 2026101 (`RECAUDO TOTAL` ASE5 = **7.550.668.665**) se reproduce **exacto** con la forma de 2 filas: `F680 = F674(0) + F648 + F613 − L613 − L648`, con `L613` = 12.379.938,06. Por tanto **NO hay insumo faltante ni cuarentena**: la fuente de ASE5 está completa y el slot `Mes2` (`F674`) vacío es correcto por diseño.

1. Entrada de divergencia **estructural** con los 6 campos Plan 42 (celda `F674` + fórmula consumidora `F680`; origen: forma-fuente de 2 filas vs geometría v2 de 3 slots; dueño: UAESP/propietaria con cita al T0; evidencia: R10 2026101 ASE5 + T0.2; fecha 2026-10-22; condición de retiro: el mapa declara slot Mes2-ASE5 o la forma-fuente trae 3 filas de modo estable — se retira la FILA, no se adapta el código).
2. Test que la pincha: `F674 == 0` + `TOT_OPT == 7.544.253.279,94` + forma-fuente citada (2 filas) + `CONSOLIDADO!D108 == 7.550.668.665` (R10 del ASE5; `D13` es su TOT_OPT = 7.544.253.279,94 — no confundir TOT_OPT con DetRetri). Si un día la fuente trae 3 filas, el test enrojece honestamente (tercera fila sin slot) en vez de absorber en silencio.
3. **Regla de filas extra (anti-silencio sin regresión):** los agregados `totOpt`/`extemp` absorben TODAS las filas por regla Σ documentada (Plan 25 D-C/D-F — caso ASE5-agosto con 3 filas Mes y mapa de 2 slots, que hoy cierra contra el R10); los slots por celda resuelven por mapa. NO se añade throw por fila-sin-slot en este camino (regresaría agosto). La red de seguridad aguas abajo (gate workbook-vs-dominio Plan 31) sigue vigente.

### D7. Usos legítimos de quincena que SOBREVIVEN (allow-list cerrada)

La quincena sigue gobernando **solo fuentes y declaración de visibles**: DetalleR2/R4 solo-Q2 (29); `AjustesSfT = 0` en Q1; columnas D/E-vs-F/G del `RESUMEN MES`; `L9:L13` + reescrituras F-NEW-1 por variante (39/40); DetRetri Q2-escribe/Q1-valida; naming por período (30); preflight saldos/retribución Q2-only (26/27); recorte conciliación-Q2 (residuo §7). El candado T4 lleva esta lista como datos (no como ramas nuevas).

### D8. Plantilla intocable + SHAs (decisión)

SHA v2 `0F192AD2…` + v1 `3086AC31…C5214` verificados antes/después de cada H (CA6). `PreparadorBasePipeline` (11 tokens sobre la COPIA) intacto; el oficial se abre read-only. Grep de cierre prueba cero diff en `EscritorValoresSlotR1.cs`, `ContratoSeccionPeriodFreeR1.cs`, `PreparadorBasePipeline*.cs` (o cita la única línea tocada — por defecto, ninguna).

### Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Reader = evidencia por firma (una sola resolución); Validador = juicio (intacto); Mapas = datos (uno productivo + uno legacy marcado); Writer = values-only (intacto) |
| **OCP** | ✅ Forma-fuente nueva (0/2/3 filas Aplic, 2/3 Mes) absorbida por la regla existente sin ramas; celda nueva = fila de mapa citada, nunca `if (quincena)` |
| **LSP/DIP** | ✅ Contratos `IRecaudoReader`/`IWorkbookLeafInputReader`/`IValidador`/`IPlantillaWriter` intactos; miembros legacy intactos; cambios detrás de las mismas abstracciones |
| **Best practices** | ✅ Contract-first (T0 congela oráculos + tablas + lista Δ antes del borrado); fail-fast core intacto en texto; tolerancia ±0.5 solo en comparación; values-only; D-H (cero literales/ramas por período o empresa — la quincena deja de ser parámetro de resolución); CRLF; sin commits; prohibido `<v>` de celda con `<f>` como oráculo |
| **Performance** | ✅ Menos código ejecutado (una resolución en vez de dos); T0 acotado BCL ZIP+XML + fixtures existentes; regens a temp fresca |

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda H)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa trx + **set-diff vs `p53-h2-suite.trx` (1289/1222/67/0)**: altas/bajas una por una con causa citada; **rojo nuevo fuera de la lista T0.6 = STOP**. 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados (re-anclajes con evidencia, §2-D7).
- G-SHA: v1 `3086AC31…C5214` + v2 `0F192AD2…` intactos antes/después de cada H.
- G-FORMULA: `<f>` byte-idéntica en todo el diff (cero `new CellFormula`, cero `InsertRow`/`DeleteRow` fuera de Phase 0).
- G-D-H: cero literales/ramas por período o empresa en producción (grep de cierre); la quincena no parametriza resolución.
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale) o la fuente ya leída; **cero workbooks extra en runtime**.
- G-TRAZA: cada error nuevo nombra ASE + reporte + fila/rol + celda donde aplique.
- G-PROC: cero commits; CRLF sin mezclar en lo tocado.

### CA — Criterios de aceptación (vinculantes; CA2/CA3 anclados al R10 oficial `R10_Remuneracion_2026101.xlsx`, verificado 2026-10-22)

- CA0: **todo número de criterio debe acompañar su derivación** (`fuente × <f>` + hoja/celda). Un número sin derivación no se acepta como criterio (lección de las dos reincidencias R-ORACULO).
- CA1: `2026101` corre los 5 ASE sin excepción.
- CA2: ASE1 **14.842.425.136** y ASE2 **11.302.637.605** = `RECAUDO TOTAL` del R10 oficial 2026101 (filas ASE 1 y 2), reproducidos por la ruta única al peso. Los valores 18.052.547.615 / 11.292.932.631 quedan **DESCARTADOS como semántica v1**: provenían del log de la corrida fallida, no de un oráculo.
- CA3: ASE3 TOT_OPT **8.811.472.160,34** (`F415 = F330+F353+F299−L299−L330`; el slot `L330` = 19.305.717,66 **sí se resta**) y EXTEMP **0,00** (`F417 = F342+F315−L315 = 0+0−0`; slots `F315`/`F342`/`L315` en el 0 del template por CERO filas `Aplicacion`). `8.830.777.878` era la fila `A=Total` de la fuente (Σ Mes, sin restar especiales) y **no** es el visible. ASE5 EXTEMP **0,00** (`F682`) y TOT_OPT **7.544.253.279,94**. El valor −1.327.632.993,07 queda eliminado como criterio; si aparece, solo citado como "valor v1 descartado por evidencia".
- CA4: `2026071` (Q1) y `2026072` (Q2) verdes **por la misma ruta** (julio-1 ASE1 con 2 filas `Mes/Total`: `Mes2 → 0m` por enmienda 2, sin cambio de valor); oráculo v2 de julio-2 (`Docs/Prueba Julio-2/Resultado/Resultado Administrativo con plantilla estandar v2/Formato Remuneracion 202607-2 Total.xlsx`) intacto.
- CA5: build con 0 warnings.
- CA6: plantilla v2 byte-idéntica salvo los 11 tokens de período aprobados.
- CA7: tras la unificación, el EXTEMP de ASE1/ASE2/ASE4 se deriva de la fórmula v2 (`Σ Aplic − L(Aplic[0])`, F6) y NO del término de subsidio v1 (F12); el T0.3 lo fija al centavo con octubre-1 (ASE1: Aplic fila 18 = 636.047,07 y fila 41 = 12.822,93; subsidio[0] fila 35 = −32.099,64 fuera del agregado).

### S1 — Criterios H0/T0 (read-only, cero código de producción)

- S1.1: geometría v2 R1 re-verificada por bytes (fronteras §0.2-F6 + textos-`<f>` de TOT_OPT/EXTEMP por ASE, esp. el de ASE3 y `F682`) + veredicto: F6 adoptado o corregido con bytes.
- S1.2: censo octubre-1 por bytes (filas `Mes/Total`, `Aplicacion`, `Total` por ASE×hoja) + fixture read-only verde que lo congela.
- S1.3: verificación CA2/CA3 (arbitraje resuelto §0.2-F13) — cada oráculo con hoja+celda+semántica citada y derivación (fuente × `<f>` v2 × semántica única) que lo reproduce al centavo; derivación ≠ criterio corregido ⇒ STOP + escalado.
- S1.4: caracterización R2/R4 Q1-vs-Q2 (valores idénticos 5 ASE × julio-1+julio-2) + celdas destino literales por bytes; cualquier Δ ⇒ STOP.
- S1.5: inventario cerrado de consumidores `EditableLeafCellsPorAse` + `F30/F10/L10` con veredicto por consumidor (re-apuntar / aislar-legacy / re-anclar-test).
- S1.6: lista cerrada de Δ esperados (los 8 E1-6-Q1 + re-anclajes + caracterización); fixture T0 verde. **Sin T0 verde no se abre H1.**

### S2 — Criterios H1 (implementación)

- S2.1: `2026101` corre 5 ASE sin excepción (CA1); ASE3 resuelve Aplic→0 por la ruta única.
- S2.2: CA2/CA3/CA7 reproducidos según §3-CA (celdas citadas por el T0.3).
- S2.3: `MapearR1`-viejo + `switch` + `ObtenerF30/ObtenerF10` eliminados (grep); línea 49 sin dispatch R1; `F30/F10/L10 = 0m` constantes (D2); `Mes2`/`Lmes2` opcionales → `0m` con test que fija julio-1 ASE1 (2 filas `Mes`) y ASE3-octubre (0 filas `Aplic`) sin aborto y sin cambio de valor (D1.1 enmendada).
- S2.4: mapa productivo único (D3); writer sin rama-R1 por quincena; R2/R4 colapsados (D4) con valores idénticos a la caracterización.
- S2.5: entrada ASE5 declarada con 6 campos + test que la pincha (D6); build 0 warnings; SHAs; G-FORMULA.
- S2.6: `R1Q2ProtectedPorAse` declara las 10 `<f>` TOT_OPT/EXTEMP (incluida `F682` de ASE5) y el test de simetría está verde (H1.7).

### S3 — Criterios H2 (regresión + candado)

- S3.1: `2026071` y `2026072` verdes **por la misma ruta**; oráculo v2 julio-2 (`Docs/Prueba Julio-2/Resultado/Resultado Administrativo con plantilla estandar v2/Formato Remuneracion 202607-2 Total.xlsx`) intacto (CA4).
- S3.2: candado T4 verde en sus dos capas (D5); set-diff == lista T0.6 exacta.
- S3.3: code-review independiente **APROBADO** (0 bloqueantes): spec compliance + calidad + greps D-H + SHAs.

### S4 — Criterios H3 (docs/cierre)

- S4.1: doctrina en `.opencode/project-context.md` (resolución R1 única; mapa único; quincena = fuentes + visibles, nunca algoritmo; candado T4; divergencia ASE5-F674).
- S4.2: este plan queda como fuente de verdad; desviaciones con fecha/motivo/evidencia; cero código de producción en H3.

---

## 4. TASKS (fases H0..H3 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero código de producción)

Fixture: `Remuneracion.IntegrationTests/ContratoUnicoLecturaR1T0Tests.cs` (BCL ZIP+XML, rutas explícitas; prohibido `Insumos.Raiz()` en lo nuevo). Evidencia: `Docs/T0-Plan54-Contrato-Unico-Lectura-R1.md`.

| ID | Tarea | Verifica |
|---|---|---|
| T0.1 | Re-verificar v2 por bytes: fronteras R1 por ASE (§0.2-F6) + adoptar la tabla de 10 `<f>` TOT_OPT/EXTEMP de F6 o corregirla con bytes (esp. `F417` ASE3 y `F682`); confirmar que el "literal 0" v1 ya no aplica | S1.1 |
| T0.2 | Censo octubre-1 por bytes: filas `Mes/Total` + `Aplicacion` + fila `Total` (valor F) por ASE; congelar §0.2-F5 como fixture (ASE3 8/24/36 + 8.830.777.878; ASE5 23/38 + 7.556.633.218; ASE1 18/41; ASE2 23/46; ASE4 31/59) | S1.2 |
| T0.3 | Verificación CA2/CA3/CA7 (arbitraje resuelto §0.2-F13): citar hoja+celda+semántica de 18.052.547.615 / 11.292.932.631 / 8.830.777.878 / 0,00 (ASE3 `F417`, ASE5 `F682`); derivar cada uno al centavo (fuente × `<f>` v2 × semántica única), incluido ASE1-octubre (F12: Aplic 18/41 dentro, subsidio[0] 35 fuera); mismatch vs §3-CA ⇒ STOP + escalado | S1.3 |
| T0.4 | Caracterización R2/R4: `MapearR2` vs `MapearR2Q2` y `MapearR4` vs `MapearR4Q2` — valores idénticos (5 ASE × julio-1+julio-2, fixtures reales) + destino literal por bytes (sin `<f>`); cualquier Δ ⇒ STOP | S1.4 (R-TABLA) |
| T0.5 | Inventario cerrado consumidores mapa-v1 + `F30/F10/L10` (writer :759/:843/:1014-1019/:1052-1108, coherencia, `ValidadorBasico.cs:812`, VerificadorRecaudo, ParidadCliTests, WorkbookLeafInputsTests, AjustesSfTTests:310, ConciliacionEmpresaTests:140) con veredicto por consumidor | S1.5 |
| T0.6 | Lista cerrada de Δ: los 8 tests Q1-E1-6 (nueva causa citada) + re-anclajes ParidadCliTests/WorkbookLeafInputsTests (con evidencia, sin tolerancia) + deltas de caracterización | S1.6 |

- **Puerta H0→H1:** fixture verde + S1.1-S1.6 cerrados. Si T0.3 o T0.4 disparan STOP, no se abre H1 (se escala al Ingeniero).

### H1/F1 — Ruta única + mapa único + colapso R2/R4 (traza a §2–§3)

| ID | Tarea | Traza |
|---|---|---|
| H1.1 | Borrar `MapearR1` + `switch (ase.Id)` + ordinales; renombrar `MapearR1Q2` → `MapearR1`; línea 49 sin dispatch R1; borrar `ObtenerF30/ObtenerF10` (previo grep de uso exclusivo); actualizar comentario HU-12/G2/G5 (40-42) con cita a este plan | D1, S2.3 |
| H1.2 | `F30/F10/L10 = 0m` constantes en la ruta única; aislar `EscribirDetalleR1` como LEGACY-only fuera de paridad; grep: cero lectores nuevos de `F30/F10/L10` | D2, S2.3 |
| H1.3 | Re-apuntar productivos a `R1Q2EditablesPorAse`; aislar `EditableLeafCellsPorAse` + helper como LEGACY (geometría v1); colapsar la rama-R1 por quincena del writer | D3, S2.4 |
| H1.4 | Colapsar R2/R4 a método + tabla únicos (SOLO si T0.4 verde); líneas 50-51 directas; borrar gemelas Q2 + mapas destino redundantes | D4, S2.4 |
| H1.5 | Entrada declarada ASE5-F674 (6 campos) + test que la pincha (`F674==0`, TOT_OPT **7.544.253.279,94** — D6/CA3 enmendados; `7.556.633.218` es la fila `A=Total` de la fuente, no el visible, forma 2-filas citada) | D6, S2.5 |
| H1.6 | Re-anclar 8 tests E1-6-Q1 + ParidadCliTests/WorkbookLeafInputsTests con evidencia T0.6 (sin tolerancia, sin Skip, sin borrados) | §2-D7, S1.6 |
| H1.7 | Cerrar el hueco de guard detectado en T0.1: `R1Q2ProtectedPorAse[5]` no declara `F682` aunque el byte la trae como `<f>F660+F629-L629</f>`; declararla y asertar la simetría de las 10 `<f>` TOT_OPT/EXTEMP | S2.6 |

- **Puerta H1→H2:** S2 verde + CA1-CA3 reproducidos + build 0w + SHAs + G-FORMULA + greps D-H.

### H2/F2 — Regresión, candado y evidencia

| ID | Tarea | Verifica |
|---|---|---|
| H2.1 | Candado T4 (estático allow-list + conductual Q1≡Q2) verde; enrojece a propósito ante `NumeroQuincena` injertado (prueba del candado) | T4, S3.2 |
| H2.2 | e2e `2026101` (CA1-CA3) + `2026071`/`2026072` misma ruta (CA4) + oráculo v2 julio-2 intacto + SHA v2 antes/después (CA6) | S3.1 |
| H2.3 | Suite trx + set-diff vs `p53-h2-suite.trx` == lista T0.6 exacta; fuera de lista = STOP | S0 G-SUITE |
| H2.4 | Code-review independiente (`code-reviewer`, este plan como spec): **APROBADO** requerido | S3.3 |

- **Puerta H2→H3:** set-diff exacto + CA1-CA7 + review aprobado.

### H3/F3 — Docs/cierre

| ID | Tarea | Verifica |
|---|---|---|
| H3.1 | Doctrina en `.opencode/project-context.md` (solo lo probado: resolución/mapa únicos, quincena = fuentes + visibles, candado, F674 declarada) | S4.1 |
| H3.2 | Cierre de este plan (§8) con desviaciones fechadas si las hubo; cero código de producción | S4.2 |

**Orden sugerido:** H0 → H1 → H2 → H3. Regresión continua (suite + SHAs en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

---

## 5. Riesgos (consolidado + gates por fase)

Ver §1.4 (R-Q1-SILENCIO, R-F30, R-TABLA, R-RECAIDA, R-ORACULO, R-ASE5, R-PLANTILLA). Riesgo residual con tratamiento propio en DESIGN: los 8 tests Q1-E1-6 cambian de causa + `ParidadCliTests`/`WorkbookLeafInputsTests` asumen `F30/F10` del mapa v1 → **re-anclaje con evidencia (T0.6/H1.6), nunca con tolerancia**: cada assert movido cita bytes; prohibido `Skip`/`Assert.True(true)`/borrados; rojo fuera de lista = STOP.

- **Gate H0→H1:** T0 verde + S1.3/S1.4 verificados sin STOP.
- **Gate H1→H2:** S2 + CA1-CA3 + 0w + SHAs + G-FORMULA + greps.
- **Gate H2→H3:** set-diff exacto + CA4-CA6 + review APROBADO.
- **Gate H3:** diff de docs citado; cero producción.

---

## 6. Clarification Gate (requiere aprobación del Ingeniero; sin esto no se abre H0)

> Resolución 2026-10-22: la pregunta 2 (oráculo CA3, EXTEMP −1.327.632.993,07) quedó RESPONDIDA por bytes v2 — EXTEMP ASE3/ASE5 = 0,00; el valor aportado era semántica v1 y queda descartado (§0.2-F12/F13, §3-CA). Se elimina del gate y se renumera.

1. **D1-D4:** ¿se ratifica borrar-vs-renombrar (U-A), `F30/F10/L10 = 0` constante + legacy aislado (L-A), mapa único con v1 aislado (D3) y colapso R2/R4 condicionado a T0.4 (D4)?
2. **ASE5-F674:** ¿la divergencia declarada entra al `DivergenciasDeclaradasCatalogo` (recomendado, con reconocimiento explícito Plan 42) o queda solo documentada en el T0 + test? ¿dueño y condición de retiro propuestos en D6?
3. **Línea base:** ¿se ratifica `p53-h2-suite.trx` (1289/1222/67/0) para el set-diff y que rojo fuera de la lista T0.6 = STOP?
4. **Enum `R1Q2Fuente`:** ¿se acepta conservar el nombre en este plan (D1.3) con el rename como follow-up?
5. **Puerta:** sin 1-4 no se abre H0; sin H1 verde no hay H2; sin review no hay H3.

Sin estas aprobaciones el plan no avanza de PROPUESTO a APROBADO, y sin APROBADO no se abre H0.

> **RATIFICADO 2026-10-22 por el Ingeniero (con calidad senior):** 1) D1-D4 aprobados (U-A, L-A, D3, D4). 2) ASE5-F674 entra al `DivergenciasDeclaradasCatalogo` con reconocimiento explícito (Plan 42), dueño UAESP/propietaria y condición de retiro de D6. 3) Línea base `p53-h2-suite.trx` (1289/1222/67/0) ratificada; rojo fuera de la lista T0.6 = STOP. 4) Enum `R1Q2Fuente` se conserva; rename como follow-up. **Enmienda 1 ratificada** (CA2/CA3/D6 anclados al R10 2026101 + CA0 + H1.7/S2.6). H0 habilitado.

---

## 7. Residuos declarados (fuera de alcance — NO implementar)

**R-RECORTE-Q2 (vivo):** `Remuneracion.Core/Services/ProcesadorPeriodo.cs:198-217` mantiene `if (!esQuincena2)` que omite la conciliación por empresa en Q2. Deuda `R-MAPA-V2-R2-ASE4` (Plan 50 §8): `Rem. Anticipos R2!E285`/`K285` es fórmula en la v2 (`E285=E192+E203-K192`). Es la otra asimetría Q1/Q2 viva después de este plan. **Evidencia:** Plan 50 §8 + T0-Plan50 §4.3. **Gate de reactivación:** T0 propio clase Plan 48 que re-congele el mapa R2 (1,4) celda-por-celda por bytes; hasta entonces el recorte permanece y el candado T4 lo lleva en allow-list documentada (no es rama de resolución). No se toca en este plan.

**R-ANCLA-EXTEMP (vivo, detectado en H2):** `Remuneracion.Core.Models.AnclasExtempR1` (Plan 49 A1-interior, medición 2026-10-19) congela `F682`/`F702` como **literales** de la v2 ("la v2 perdió la fórmula del manual"), pero los bytes de la v2 **vigente** (SHA `0F192AD2…7F2DAB`) traen `F682 = <f>F660+F629-L629</f>`. `ValidacionOracleReader.ResolverAnclasExtempR1` (que exige literal) aborta con `ERR-PLANTILLA` en toda corrida 5-ASE — **agnóstico de quincena**: muerde Q1 y Q2 e igual los tres `ParidadCliTests` y 5 tests `DivergenciaDeclarada*`. **Estado:** rojo de **BASE**, byte-idéntico antes y después del Plan 54 (`AnclasExtempR1.cs` **no fue tocado**; el rojo ya estaba en `p53-h2-suite.trx`). **Gate de reactivación:** T0 de bytes que re-congele las anclas contra la v2 vigente y plan propio (clase Plan 48/49). **No se toca acá.**

**R-CIERRE-EMPRESAS-2026101 (vivo, detectado en la verificación E2E real post-cierre):** con la ruta única, la corrida CLI `2026101` **pasa la lectura de los 5 ASE** y los 5 `DetRetri` esperados coinciden con el R10 oficial al peso, pero aborta más adelante en `WorkbookLeafCoherence.ValidarSigmaEmpresas` (`OpenXmlPlantillaWriter.cs:236-239`; regla: solo Q1, `NumeroQuincena != 1 ⇒ return`) porque **Σ visibles por empresa ≠ visible de bloque**: ASE1 R1 `14.815.023.989,82` vs `14.815.021.974,75` (Δ `2.015,07` = |primera fila `Mes`|); ASE4 R1 Δ `309.350,00` = `ReciprocidadEaab`; ASE2 R2 Δ `130.524,00`; ASE1 R4 Δ `3.698.870,00`; ASE2 R4 Δ `945.063,00`. **No lo causa el Plan 54**: el agregado por empresa y los mapas R2/R4 no se tocaron; y con la semántica v1 de ASE1 (2 de las 3 filas `Mes`) el Δ era de orden miles de millones, no de `2.015,07` — el plan **reduce** la brecha, no la crea. Es la **clase del agregado por empresa** (`R-MAPA-V2`/Plan 50 A3) y **no está cubierta por ningún test** (0 ocurrencias del mensaje en `p53-h2-suite.trx`), por eso no aparecía en la línea base. **Gate de reactivación:** plan propio (Plan 55) con T0 de bytes que re-congele el mapa por empresa R1/R2/R4 contra la v2 y explique cada Δ por celda; hasta entonces la corrida 5-ASE de Q1 no cierra la validación. **No se toca acá.**

---

## 8. Estado

**CERRADO 2026-10-22.** Aprobado por el Ingeniero (Clarification Gate §6 ratificado + enmiendas 1 y 2), implementado en H0→H3 y **aprobado por code-review independiente con observaciones (0 bloqueantes)**.

**Evidencia final (H3.2, `TestResults/p54-H3-suite.trx`; idéntica a `p54-H2-suite-final.trx`):**

| Gate | Resultado |
|---|---|
| G-BUILD | `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores** |
| G-SUITE | **1374 / 1307 / 67 / 0** vs base `p53-h2-suite.trx` (1289/1222/67/0): **+85 altas verdes, 0 bajas, 67 rojos = MISMO conjunto (0 nuevos, 0 desaparecidos)** |
| G-SHA | v2 `0F192AD2…7F2DAB` y v1 `3086AC31…C5214` **idénticos** antes/después de cada fase |
| G-FORMULA | cero `new CellFormula`, cero `InsertRow`/`DeleteRow`; `<f>` byte-idéntica |
| G-PROC | 0 commits; CRLF puro; `EscritorValoresSlotR1`, `ContratoSeccionPeriodFreeR1`, `PreparadorBasePipeline` y `ProcesadorPeriodo.cs:198-217` sin tocar |

**Trabajo cerrado:** T1 (una sola resolución R1: `MapearR1` por firma, `Mes0/Mes1` core, `Mes2/Lmes2` y `Aplic*` opcionales → 0; `ObtenerF30/F10` y el `switch (ase.Id)` eliminados) · T2 (mapa productivo único `R1Q2EditablesPorAse`; `WorkbookLeafCellMapPorAse` LEGACY) · T3 (R2/R4 a tabla única) · T4 (candado de dos capas) · H1.7/S2.6 (`F682` declarada en `R1Q2ProtectedPorAse[5]` + simetría de las 10 `<f>`).

**Desviaciones fechadas:**

- **(a) Enmienda 1 (2026-10-22)** — CA2/CA3/D6 re-anclados al R10 oficial de octubre-1, tras **STOP del implementer en T0.3**: los criterios aportados (`18.052.547.615` / `11.292.932.631` / `8.830.777.878` / `7.556.633.218`) no eran derivables de la fuente por la ruta v2; se reconstruyeron como semántica v1 (la fila `A=Total` de la fuente es `Σ Mes + Σ Aplic` y **no resta Especiales**). Se añadió **CA0** (todo criterio numérico viaja con su derivación `fuente × <f>` + hoja/celda).
- **(b) Enmienda 2 (2026-10-22)** — `Mes2`/`Lmes2` **opcionales → `0m`**, tras **Δ-1 del T0**: julio-1 ASE1 trae 2 filas `Mes/Total` y el mapa v2 le declara `Mes2`; con `Mes2` core la ruta única abortaba y **CA4 era inalcanzable**. Con `Mes2 → 0` reproduce su R10 al peso (`16.758.167.585`) **sin cambio de valor**.
- **(c) Límite de CA4 (2026-10-22)** — **no se aserta `CodigosSalida.Ok`**: los tres `ParidadCliTests` son **rojos de base** por `ERR-SIN-SLOT` (contrato de sección, Planes 41/46) y `ERR-PLANTILLA` (`R-ANCLA-EXTEMP`, §7). CA4 se asertó en la **ruta**: identidad Q1≡Q2 sobre la misma fuente, lectura de los 5 ASE en los 3 períodos (`2026101`, `2026071`, `2026072`) y ausencia del mensaje del defecto.
- **(d) O1 del code-review (2026-10-22)** — la allow-list de **D5** se lee **subsumida en D7** (lista normativa de usos legítimos de la quincena); el candado implementa exactamente los tres usos de D7 (`DetalleR2/R4` solo-Q2, D/E-vs-F/G del RESUMEN MES, dato de procedencia `Quincena =`).

**Correcciones de oráculo registradas (para que no se repitan):**

1. `−1.327.632.993,07` (EXTEMP ASE3) → **descartado**: era semántica v1 (`F243+F223−L223`, con término de subsidio); los bytes v2 dan `F417 = F342+F315−L315 = 0,00`.
2. `18.052.547.615` / `11.292.932.631` / `8.830.777.878` / `7.556.633.218` → **descartados**: los dos primeros eran el resultado de la rama v1 (log de la corrida fallida) y los dos últimos la fila `A=Total` de la fuente. Criterios vigentes: `14.842.425.136` / `11.302.637.605` / `8.811.472.160,34` / `7.544.253.279,94`, todos con derivación (CA0).

**Frase de cierre:** *la lectura R1 tiene un solo camino para ambas quincenas; la quincena ya no selecciona algoritmo.*

**Verificación E2E real post-cierre (2026-10-22, posterior a H3):** `dotnet run --project Remuneracion.Cli -- --periodo 2026101 --carpeta "Docs/Prueba Octubre-1/Insumos" --plantilla "<v2>" --salida <tmp> --cinco-ase --sobrescribir`:
- **El defecto del plan ya no ocurre**: los 5 ASE completan la lectura; desaparece `"no se encontró la fila 'Aplicacion nuevos x reversion'"`.
- **`DetRetri esperado post-Excel` = R10 oficial en los 5 ASE, al peso**: `14842425136` / `11302637605` / `8823787300` / `10483172762` / `7550668665`.
- La corrida **aún no cierra**: aborta después en `ValidarSigmaEmpresas` por **`R-CIERRE-EMPRESAS-2026101`** (§7), defecto de la clase agregado-por-empresa, ajeno a este plan y sin cobertura de test. **Siguiente paso propuesto: Plan 55.**
