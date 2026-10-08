# Plan 31 — T4 Veredictos R2/R4-remesh + hojas de soporte

> Alcance: decisión documentada por frente (R-D-1..R-D-6, más el caso condicional
> ANT EXT-REV / ANTICIPOS USUARIOS) para el remesh R2/R4 y las hojas de soporte, con consecuencia
> explícita por frente (comparador: re-encender o versionar como brecha declarada; código:
> follow-up con T0 propio o nada). Este documento NO abre producción. El único cambio de código
> admitido por T4 es una rama de exclusión versionada en el comparador (R-D-3 / R-D-4 / R-D-6),
> acotada por hoja y con cita a este documento; el resto del comparador y toda la producción quedan
> intactos.
>
> Fuente única de verdad (rector): `plans/31 - recomposicion-totales-R1-por-firma.md` (§0.3 D-G,
> §1.3 Non-Goals, §2.3 contrato T2, §3.1 R-D-1..R-D-6, §4 T4); evidencia T0b
> `plans/29-T0-Evidencia-B.md`; evidencia R1 `plans/31-T0-Evidencia-R1.md`; `plans/25`
> (roles R1 por firma), `plans/23` (HU-22, opcionalidad `Vlr Intereses`), `plans/11` (HU-11
> SALDOS POR NOTA / RETRIBUCION NEGATIVA).
>
> Método: re-lectura de T0b §1..§2 + medición EN VIVO del rojo del comparador
> (`ComparadorSalidaVsManualTests`, BCL puro ZIP+XML, sin Excel/COM) + lectura de los cell-maps y
> del mutador. Toda cifra remite a la corrida medida del 2026-10-07 o a archivo + hoja + celda.
> Sin datos inventados.
>
> Estado de entrada: T1/T2/T3 cerrados (sin commit). Julio identidad; el gate R1
> (`ValidadorTotalesR1GateTests`) verde en ambos períodos para los 3 visibles del contrato.
> Fecha: 2026-10-07.

---

## 0. Tabla de veredictos

Numeración: el encargo enumera 7 frentes candidatos para los 6 IDs R-D-1..R-D-6; este documento
mapea R-D-1..R-D-6 a los 6 frentes con rojo medido y agrega el séptimo (ANT EXT-REV / ANTICIPOS
USUARIOS, que "si aparece") como R-D-7 condicional. `DetRetri*/DetValiRetri* E..O por cadena`
(que el plan §3.1 listaba) queda como nota: ya está versionado por el comparador (ítem 1,
`Unidad D (T4/T0d)`), no reabre.

| ID | Frente | Veredicto | Consecuencia (comparador / código) | Cita |
|---|---|---|---|---|
| R-D-1 | R2-detail (hoja `Rem. Anticipos R2`) | declarar-divergencia | Comparador: YA versionado (brecha ítem 9). Código: nada (el mapeo por encabezado/label ya es el elegido; adoptar la geometría del manual exige insertar/borrar filas, prohibido hoy) | T0b §2.1, §2.2, §2.4 |
| R-D-2 | R4-detail (hoja `Reversion Pagos R4`) | declarar-divergencia | Comparador: YA versionado (brecha ítem 9). Código: nada (idem R-D-1) | T0b §2.3, §2.4 |
| R-D-3 | SALDOS POR NOTA (detalle) | declarar-divergencia | Comparador: versionar la categoría (ramas nuevas, agosto). Código: nada (doctrina Plan 23 "ausente = 0 explícito"); re-encender cuando se cierre la malla (F-T4-2/F-T4-3) | T0b §1.3/§2.4; `WorkbookLeafCellMapAjustesSfT`; Plan 23 (HU-22) |
| R-D-4 | AJUSTES - SF-T + arrastre (`CONSOLIDADO_TOTAL RECAUDO`, `REMUNERACION_*`) | declarar-divergencia | Comparador: versionar la categoría (ramas nuevas, agosto). Código: nada (arrastre de R-D-1/R-D-2/R-D-3); re-encender al cerrar la malla | T0b §1.3/§2.4; `WorkbookLeafCellMapAjustesSfT.Protegidas` |
| R-D-5 | R1-interior subvisible (`Reporte Componentes R1`, F/G/H interiores) | adoptar-geometría (follow-up con T0 propio) | Comparador: NO versionar; queda en ROJO como brecha viva (S3 "cero divergencias R1 fuera de versionadas" NO se cumple aún). Código: FOLLOW-UP F-T4-1 (extender el contrato de recomposición al interior) | T0 §1.6; `TotalesR1Esperados` (nota); `ValidadorTotalesR1Workbook` (contrato) |
| R-D-6 | INTERVENTORIA R26 | declarar-divergencia (cosmética) | Comparador: versionar (acotado a `R26`, agosto). Código: nada (mismo valor; edición de fórmula del manual) | Medición 2026-10-07: `INTERVENTORIA!R26` app `=F15` vs manual `=H15`, valor 189185988 en ambos |
| R-D-7 | ANT EXT-REV / ANTICIPOS USUARIOS (condicional) | fuera-de-alcance | Comparador: YA versionado (brecha ítem 8, non-goal del Plan 29). Código: nada | T0b §0/§1.3; `ComparadorSalidaVsManualTests.BrechaDeclarada` ítem 8 |

Resumen del estado operativo del comparador tras R-D-3/R-D-4/R-D-6: agosto pasa de 1272 a 96
divergencias inesperadas, todas en `Reporte Componentes R1` (el frente R-D-5, follow-up). Julio
sigue en 0.

---

## 1. Evidencia por frente

### 1.0 Medición base del rojo (2026-10-07)

Corrida en vivo de `ComparadorSalidaVsManualTests` (ambos períodos, insumos reales, salida
temporal; no toca `Docs/`):

- julio (`202607-2`): 0 divergencias inesperadas (PASS).
- agosto (`202608-2`): **1272 divergencias inesperadas de 6275** celdas comparadas.

Desglose por hoja (líneas deduplicadas; la consola repite cada divergencia por el logger):

| Hoja | Divergencias | Familia |
|---|---:|---|
| `SALDOS POR NOTA` | 313 | malla de detalle (R-D-3) |
| `CONSOLIDADO_TOTAL RECAUDO` | 140 | arrastre (R-D-4) |
| `REMUNERACION_ENERBIT` | 135 | arrastre (R-D-4) |
| `REMUNERACION_ENEL` | 135 | arrastre (R-D-4) |
| `REMUNERACION_OCCIDENTE_ Directa` | 135 | arrastre (R-D-4) |
| `REMUNERACION Reciprocidad EAAB` | 135 | arrastre (R-D-4) |
| `REMUNERACION_EAAB-CL` | 126 | arrastre (R-D-4) |
| `Reporte Componentes R1` | 96 | interior subvisible (R-D-5) |
| `AJUSTES - SF-T` | 56 | arrastre de SALDOS (R-D-4) |
| `INTERVENTORIA` | 1 | cosmética (R-D-6) |
| **Total** | **1272** | |

Todas son divergencias de TEXTO de fórmula por desplazamiento de filas ("misma familia mesh", Plan
31 §1.6 de la evidencia R1); no son diferencias de VALORES de dominio. Las hojas `Rem. Anticipos R2`
y `Reversion Pagos R4` NO aparecen en la lista porque ya están versionadas por el comparador
(ítem 9, agosto).

### 1.1 R-D-1 — R2-detail (hoja `Rem. Anticipos R2`)

- Evidencia T0b (`plans/29-T0-Evidencia-B.md`):
  - §1.3 / §2.1: las columnas de la fuente R2 derivan entre períodos en TODOS los ASE (julio
    E..P sin `Especiales`; agosto E..Q con `Especiales` en K). El corrimiento posicional fijo de
    D-C queda refutado; el mapeo robusto es por NOMBRE de encabezado.
  - §2.2: las filas derivan en `LIME`-agosto (46->47) y `BOGOTA LIMPIA`-agosto (49->50); los otros
    3 ASE coinciden. Plantilla == julio en 5/5.
  - §2.4 (`RECORTE R-DESGOSE-DERIVA`): agosto R2 estable en {1,3,5}; prohibido insertar/borrar
    filas (D-C).
- Código: la app mapea el detalle por encabezado/label a las filas canónicas de la plantilla
  (`WorkbookLeafCellMapDetalleR2R4`). El manual del administrativo recompone la malla (intercala
  sub-bloques). El comparador compara celda-a-celda contra la geometría del manual.
- Veredicto: **declarar-divergencia**. La app sigue la geometría canónica de la plantilla; el
  manual recompone. Ya está versionado (brecha ítem 9: `hoja is "REM. ANTICIPOS R2" or
  "REVERSION PAGOS R4"` cuando `esAgosto`). No se toca.

### 1.2 R-D-2 — R4-detail (hoja `Reversion Pagos R4`)

- Evidencia T0b §1.3 / §2.3: R4 fuente julio 13/13/13/19/13 == plantilla 5/5; agosto 6/13/7/18/13
  -> sólo `LIME` (ASE2) y `AREA LIMPIA` (ASE5) coinciden. Offset destino = fuente − 2. Filas de
  cierre con fórmula protegida (`r73=D15-P15`, etc.).
- §2.4: agosto R4 estable en {2,5}; prohibido insertar/borrar filas.
- Veredicto: **declarar-divergencia** (idem R-D-1). Ya versionado (brecha ítem 9). No se toca.

### 1.3 R-D-3 — SALDOS POR NOTA (detalle)

- Medición: 313 divergencias en agosto (la mayor de las hojas de soporte).
- Lectura de código (`WorkbookLeafCellMapAjustesSfT`):
  - `FilasSaldosNotasPorAse`: filas editable = canónicas por ASE (1: 3..7; 2: 15..20; 3: 28..32;
    4: 40..45; 5: 53..58).
  - `ConceptosSaldosPorAse`: mapeo fuente->template POR TÍTULO de concepto (prohibidos offsets,
    T0-0.7). `Vlr Intereses` es opcional (Plan 23 / HU-22): ausente en la fuente = 0 explícito, la
    fila canónica se conserva.
- Evidencia medida (muestras de la corrida):
  - `SALDOS POR NOTA!B16` app `Vlr Intereses` vs manual `Total`.
  - `SALDOS POR NOTA!C16` app `0` vs manual `931073.72`; `C17` app `931073.72` vs manual
    `-424218.36` (desplazamiento de fila).
  - `SALDOS POR NOTA!C22` app `=f(C20-I20)` vs manual `=f(C18)` (visible `Cn-In` de cada geometría,
    consistente en ambas).
- Interpretación: la app escribe la geometría CANÓNICA de la plantilla (conserva la fila opcional
  como 0, doctrina Plan 23); el manual del administrativo recomputa la malla (elimina la fila
  opcional ausente y re-fluye). Misma familia que R2/R4 (T0b §2.4): las celdas visibles de cada
  libro son internamente consistentes; sólo difiere el TEXTO de las referencias. No hay defecto de
  valor.
- Veredicto: **declarar-divergencia**. Consecuencia comparador: versionar la categoría (rama nueva,
  agosto). Código: nada. Se re-enciende cuando la malla se unifique (F-T4-2/F-T4-3) o se retire la
  exclusión si la decisión es permanente.

### 1.4 R-D-4 — AJUSTES - SF-T + arrastre (`CONSOLIDADO_TOTAL RECAUDO`, `REMUNERACION_*`)

- Medición: `AJUSTES - SF-T` 56; `CONSOLIDADO_TOTAL RECAUDO` 140; `REMUNERACION_*` 666
  (135 x 4 + 126).
- Evidencia medida (muestras):
  - `AJUSTES - SF-T!D10` app `=f('SALDOS POR NOTA'!C22)` vs manual `=f('SALDOS POR NOTA'!C21)`;
    `AA10` app `=f('SALDOS POR NOTA'!O22)` vs manual `=f('SALDOS POR NOTA'!O21)`.
  - `CONSOLIDADO_TOTAL RECAUDO!AA29` app `=f('Rem. Anticipos R2'!Q139)` vs manual `Q140`;
    `AA66` app `=f('Reversion Pagos R4'!O73)` vs manual `O66`.
  - `REMUNERACION_OCCIDENTE_ Directa!F29` app `=f('Rem. Anticipos R2'!G151)` vs manual `G152`;
    `F66` app `=f('Reversion Pagos R4'!F85)` vs manual `F78`.
- Cadena documentada por el cell-map (`WorkbookLeafCellMapAjustesSfT.Protegidas`): los visibles de
  `AJUSTES - SF-T` (`D9:D13`, `D28:D32`, `D47:D51`) referencian los visibles `Cn-In` de
  `SALDOS POR NOTA` / `RETRIBUCION NEGATIVA`; `CONSOLIDADO D85:D89` referencia `AJUSTES D47:D51`;
  y `REMUNERACION_*` referencia el detalle R2/R4.
- Interpretación: todas son fórmulas que referencian una malla recomputada con filas desplazadas;
  el TEXTO difiere por el desplazamiento de fila, no el valor de dominio (la app referencia SU
  propia geometría, internamente consistente). Puro arrastre de R-D-1/R-D-2/R-D-3.
- Veredicto: **declarar-divergencia** (arrastre). Consecuencia comparador: versionar la categoría
  (`AJUSTES - SF-T`, `CONSOLIDADO_TOTAL RECAUDO`, prefijo `REMUNERACION`, agosto). Código: nada.

### 1.5 R-D-5 — R1-interior subvisible (hoja `Reporte Componentes R1`)

- Medición: 96 divergencias en agosto; todas en `Reporte Componentes R1` (columnas F/G/H, filas
  INTERIORES de cada bloque, no los 3 visibles del contrato).
- Evidencia de contrato (por qué quedaron fuera de T2):
  - El contrato T2 (Plan 31 §2.3) y el gate `ValidadorTotalesR1Workbook` cubren SOLO tres visibles
    por ASE: TOT_OPT F, total TDF G y EXTEMP F (`ValidarAse`, `TotalesR1Esperados`).
  - `TotalesR1Esperados` lo declara explícitamente: "Sub-visibles de empresa (RECIPROCIDAD/ENEL/…):
    el gate T1 no los consume (solo valida TOT_OPT, TDF y EXTEMP del bloque), por eso no se
    congelan aquí".
- Evidencia medida (muestras):
  - `Reporte Componentes R1!F60` app `=f(F40+F21+F8-L8-L21)` (3 términos) vs manual
    `=f(F37+F18-L18)` (2 términos): la app conservó la FORMA de julio, no la geometría recomputada
    de agosto.
  - `F51` app `=f(F47)` vs manual `=f(F45)`; `F217` app `=f(F138)` vs manual `0`; `F464` app `0`
    vs manual `=f(F414)`.
- Interpretación: NO es el caso "app canónica vs manual recompute" de R-D-1..R-D-4. El interior de
  R1 lo reescribe el reanclaje MECÁNICO del mutador (`Reanclar`, mapa `primeraBorrada` calibrado a
  julio; evidencia R1 §1.6), el mismo mecanismo cuya falla T2 corrigió SÓLO para los 3 visibles del
  contrato. En agosto (fuente que recorta en la CABEZA) las anclas interiores quedan desplazadas ->
  DEFECTO de anclaje real, no divergencia de representación.
- Veredicto: **adoptar-geometría (follow-up con T0 propio)**. Consecuencia:
  - Comparador: NO se versiona. Estos 96 quedan en ROJO como brecha viva; la promesa S3 ("cero
    divergencias R1 fuera de versionadas" post-T2) NO se cumple por completo — sólo se cumple para
    los 3 visibles del contrato. Esto se reporta como la brecha abierta de T4.
  - Código: FOLLOW-UP F-T4-1 (ver §2). Versionar estos 96 ocultaría un defecto real, por eso se
    descarta.

### 1.6 R-D-6 — INTERVENTORIA R26

- Medición: 1 divergencia en agosto: `INTERVENTORIA!R26` app `=f(F15)` vs manual `=f(H15)`.
- Verificación de valor (2026-10-07, lectura directa de los dos workbooks, rango `INTERVENTORIA!F13:R26`):
  las matrices leídas de la salida de la app y del manual son IDÉNTICAS en ese rango; en particular
  `F15 = H15 = 189185988` y `R26 = 189185988` en ambos. La diferencia es SÓLO de texto de fórmula.
- Lectura de código: `WorkbookLeafCellMapAjustesSfT` documenta las fórmulas protegidas de
  `INTERVENTORIA` (`K31`/`M31`/`N31`, `SUM`) y la zona `L25:N31` como no escrita; `R26` no está en
  el mapa de escritura -> la app conserva la fórmula de la plantilla. La divergencia es una edición
  cosmética de la fórmula del manual (misma referencia de valor, distinta celda fuente).
- Veredicto: **declarar-divergencia (cosmética)**. Consecuencia comparador: versionar la categoría,
  acotada a `d.Celda == "R26"` en agosto. Código: nada.

### 1.7 R-D-7 — ANT EXT-REV / ANTICIPOS USUARIOS (condicional)

- No aparecen en el rojo medido: el comparador ya los versiona (ítem 8: `hoja is "ANT EXT-REV" or
  "ANTICIPOS USUARIOS"` -> "hoja no escrita por la app (non-goal del Plan 29)"). T0b confirma que
  `ANT EXT-REV` es zona protegida (no escrita por el writer Q2).
- Veredicto: **fuera-de-alcance** (non-goal del Plan 29, ya operativo). Sin cambio.

### 1.8 Nota — `DetRetri*`/`DetValiRetri* E..O por cadena` (plan §3.1)

Ya versionado por el comparador (ítem 1, `Unidad D (T4/T0d)`): la app escribe `DetRetri D` y
`DetValiRetri I/O` por ASE; el resto es recorte declarado por origen fórmula-cadena o manual-externo
(`DetValiRetri!J`, SALE de T0d §1.4). No reabre en T4.

---

## 2. Follow-ups (numerados)

- **F-T4-1 (R-D-5, adoptar-geometría):** Plan 32 — extender la recomposición por firma al INTERIOR
  de `Reporte Componentes R1` (sub-visibles de empresa y totales interiores F/G/H). Requiere T0
  propio: congelar E2 del texto esperado por ASE/período del manual (solo lectura ZIP+XML), ampliar
  el contrato de `OpenXmlEspejoR1Mutador` (pase final) y de `ValidadorTotalesR1Workbook` +
  `TotalesR1Esperados`. Verificación de cierre: comparador R1 = 0 inesperadas fuera de versionadas.
- **F-T4-2 (R-D-1/R-D-2, adoptar-geometría del remesh R2/R4):** sólo si el Ingeniero AUTORIZA el
  mapeo dinámico de FILAS (label A/B/C/D) o insertar/borrar filas (hoy prohibido por D-C / T0b
  §2.4). T0 propio sobre `Rem. Anticipos R2` / `Reversion Pagos R4`. Si se implementa, re-encender
  la brecha ítem 9 (retirarla).
- **F-T4-3 (R-D-3/R-D-4, re-encender):** `SALDOS POR NOTA`, `AJUSTES - SF-T`,
  `CONSOLIDADO_TOTAL RECAUDO` y `REMUNERACION_*` vuelven a compararse si F-T4-2 unifica la malla; si
  la decisión es mantener la geometría canónica como permanente, promover estas exclusiones a
  "exclusión versionada permanente" (producción) con su propia cita, en vez de brecha temporal.
- **F-T4-4 (R-D-6):** si el Ingeniero confirma que `INTERVENTORIA!R26` es del manual (cosmética), la
  exclusión permanece; si la app debe gobernarla, corregir la referencia en el mapa de
  INTERVENTORIA.

---

## 3. Cambio de código aplicado por T4 (único admitido)

Rama de exclusión versionada en `BrechaDeclarada` de
`Remuneracion.IntegrationTests/ComparadorSalidaVsManualTests.cs`, acotada a agosto y a las hojas
cuya categoría es "declarar-divergencia" (R-D-3, R-D-4, R-D-6), con cita a este documento. No se
modifica ninguna otra lógica del test. Las categorías YA versionadas (R-D-1/R-D-2 ítem 9, R-D-7
ítem 8, Unidad D ítem 1) no se tocan.

Consecuencia medida: agosto pasa de 1272 a 96 inesperadas (todas R-D-5, brecha viva); julio sigue en
0. Estas exclusiones son temporales y deben retirarse cuando se cierren F-T4-1/F-T4-2/F-T4-3.
