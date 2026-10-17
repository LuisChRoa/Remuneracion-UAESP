---
titulo: Plan 46 - Correccion mapeo interior R1 (canal efectivo + seccion por cierre)
estado: CERRADO 2026-10-17 — H0 (T0 corregido read-only) + H1 (sección por CIERRE + canal efectivo) + H2 (fixtures + regresión + re-corrida; opción A: retiro de la entrada OCCIDENTE-ASE5 + fixture sintético) + H3 (docs) + review W-1/W-2 cerrados. Set-diff final 0 altas / 1 baja documentada; build 0 warnings; oficial SHA256 intacto.
fecha: 2026-10-17
numero: 46
planes_ocupados: 01..45 ocupados; este es el 46
padre: plans/45 - contrato-seccion-periodfree-sinc.md (CERRADO 2026-10-16)
abuelo: plans/44 - llenado-r1-fuente-canal-concepto-componentes.md (F1 DONE + F2/F3 y paridad-auditoria PENDIENTES)
gate_entrada:
  decision_ingeniero_2026-10-17: corregir el mapeo del interior de `Reporte Componentes R1` con los 2 defectos verificados celula a celula en la corrida real del App (2026-10-17, salida `Docs/Prueba Julio-2/Resultado/8/Remuneracion 202607-2 Total.xlsx`)
  q1: NO SE CAMBIA comportamiento sin evidencia; Q1 solo se MIDE en la re-corrida (formas [A,M,A,M] comparten regla con Q1-ASE1/5; neutralidad Q1 del Plan 45 se preserva; `SeccionEspejoR1Tests` con Q1 sigue verde)
  plantilla_oficial: NO SE TOCA (SHA256 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` intacto antes/despues)
  catalogo_divergencias: NO SE TOCA (`DivergenciasDeclaradasCatalogo` intacto; Via A/B solo recibe numeros frescos post-fix, no se decide aqui)
  alcance: contrato de seccion corregido + canal efectivo + erratas T0 + retiro de 2 ceros documentados + fixtures actualizados con evidencia + re-corrida completa con evidencia
---

# Plan 46 - Correccion mapeo interior R1 (canal efectivo + seccion por cierre)

> SDD completo PROPOSE -> DESIGN -> SPEC -> TASKS. Fuente unica de verdad para la implementacion una vez aprobado ESTE documento. No se implementa sin aprobacion del Ingeniero.
> Lectura obligatoria previa: plan padre completo (45) + Enmienda E-44.1 (plan 44 §20) + `Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md` (§3, §4, §5, §9 — secciones con errata, ver §6 de este plan) + `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs` (`ClasificarSeccionesRaw`, `SeccionQueAbre`, `ClasificarSeccionesContrato`, `EncajaEnSlot`, `ConstruirGruposSeccion`, `BuscarGrupo`, `TryPlanificar`, `LeerSlots`, `Escribir`) + `Remuneracion.Core/Models/ContratoSeccionPeriodFreeR1.cs` (`Alinear`, `ParticipaDeSeccionConformada`) + `Remuneracion.Core/Models/SeccionEspejoR1.cs` (consumidor Plan 35-T2 intacto) + `Remuneracion.Core/Models/CerosDocumentadosR1.cs` + `Remuneracion.Core/Models/R1FirmaInterior.cs` (`EsDatoEmpresa` — baseline Plan 41, intacto) + `.opencode/project-context.md` (doctrinas e invariantes, § Plan 45) + `Docs/Instructivo-Capa-B.md` si aplica.
> Visual Design Intent: N/A (cero UI en este plan).

## 0. Contexto adoptado (no re-litigar; H0 lo re-deriva contra bytes)

### 0.1 Origen: corrida real del App 2026-10-17 (oraculo de este plan)

- La corrida real del App (salida `Docs/Prueba Julio-2/Resultado/8/Remuneración 202607-2 Total.xlsx`) demostro **2 defectos** en el interior de `Reporte Componentes R1`, verificados **celula a celula por el orquestador con lectura binaria ZIP+XML** (helper de volcado disponible en `C:\Users\lechaparro.ASEINGES\AppData\Local\Temp\opencode\dumpxlsx.ps1` con params `-Path/-Sheet/-From/-To/-Cols/-List`).
- Medicion dev (NO oraculo, NO relacionada con la ejecucion del App): `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx` — referencia de desarrollo; la medicion admin pone O-VS y O-Total en su par O (evidencia de contraste, no spec).
- Insumos (fijos de este plan): fuente `Docs/Prueba Julio-2/Insumos/{1-Promoambiental,2-Lime,3-Ciudad Limpia,4-Bogota Limpia,5-Área Limpia}/Recaudoporcomponente_*.xlsx` (una sola hoja `Sheet1`; ojo: en `5-Área Limpia` hay `.pdf` homonimos — usar el `.xlsx`). Oficial: `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` (hoja `Reporte Componentes R1`).

### 0.2 Anclas de geometria oficial (adoptadas, H0 las re-deriva)

- ASE1 cabecera R4, bloque 6-72, fronteras Mes/Aplic en 16/32/47/59/70.
- ASE5 cabecera R601, bloque 603-676, fronteras 613/629/648/660/674.
- Secciones oficiales ASE5: P=(603-613, cerrada por Mes@613), Aplic#1=(614-629, cerrada por Aplic@629; data rows 615-628: E 615-618, H 619-621, O 622-624, T 625-627, C=OCC 628), Oportuno#2=(630-648), Aplic#2=(649-660), Oportuno#3=(661-674), cola 675-676.

### 0.3 Falsas pistas (resueltas — se documentan para que no reaparezcan)

| # | Falsa pista | Veredicto (no reabrir) |
|---|---|---|
| FP-1 | `370,08` / `-1,95` / `F649=13.359.789.539,23`: residen DENTRO de la plantilla oficial (12 celdas en sheet10). | La App no los escribe; donde hay dato se sobrescriben, donde no quedan. NO tocar la plantilla. Nota opcional para el dueno (sin enviar). |
| FP-2 | "Componentes de OCCIDENTE en 0" en bloque inferior ASE1 (R30/R31/R58). | Correcto-por-fuente: el bloque Aplicacion de ASE1 solo trae ENEL E; el sub-bloque de subsidios no trae H/O/T. Sin accion. |
| FP-3 | "Hueco de plantilla Vlr Intereses" del T0 §5/§9 (402,40 y 7.697,23 de ASE5 1.ª). | Con la seccion corregida SI tienen slot (F623/F626). WITHDRAW del texto §9 (no se reporta al dueno) y retiro de ambas entradas de `CerosDocumentadosR1`. El mecanismo de cero documentado se conserva vacio/sin esas filas. |

### 0.4 Linea base y puertas heredadas

- Linea base de suite a la entrada: `Remuneracion.IntegrationTests/TestResults/o43-final-failed.txt` (**39 fallos**; referencia `p45-h2b.trx` 968/929/39).
- Puertas del padre/abuelo intactas y referenciadas, no reabiertas: auditoria de la propietaria PENDIENTE (no se declara paridad), Via A/B solo recibe numeros frescos (no se decide aqui), otras pestanas fuera, plantilla oficial intacta, cero commits.

---

## 1. PROPOSE

### 1.1 Problema (2 defectos probados en la corrida 8)

**DEFECTO 1 — corrimiento de canal en filas sin D (VS/VI/Subsidio).** Las filas sin canal propio (D vacio: VS, VI, Subsidio) se asignan por **cursor de orden dentro de la (firma, seccion)** en vez de por su canal. Evidencia corrida 8:

- ASE1 Mes: O-VS -4.108,55 quedo en F9 (fila H; deberia ser F11) y F11 quedo vacio; F12 (O-Total) correcto. La medicion admin pone O-VS y O-Total en su par O.
- ASE5: H-VS 52.706,16 -> F603 (fila E, deberia F619); O-VS 49.921,06 -> F606 (deberia F622); T-VS 306.792,77 -> F608 (deberia F625); los totales H/O/T si cayeron bien (F607/F609/F611 -> van a F621/F624/F627). Subsidios: H -7.906,16 -> F650 (deberia F653), O 32.235,54 -> F653 (deberia F655).

**DEFECTO 2 — seccion de la 1.ª seccion fuente mal tipada.** `ContratoSeccionPeriodFreeR1` tipa la 1.ª seccion fuente como `Primera` y la alinea a la 1.ª oficial. Correcto: **cada seccion se tipa por la frontera que la CIERRA** (la inmediatamente debajo: su fila resumen `Mes`/`Aplicacion…` cuyo importe es su subtotal). ASE5-julio2: la 1.ª (filas 5-13) la cierra Aplicacion (F417.519,62 en su fila 14) -> seccion Aplicacion -> cae en Aplic#1 oficial (615-628). Con esto: 603-614 queda en cero ("no vino nada de AFaseo") y 615-628 recibe el bloque.

Formas resultantes:

- `[M,A,M,A,M]` → identidad (sin cambio).
- `[A,M,A,M]` (ASE5-julio2, Q1-ASE1/5) → `[A#1(615-628), O#2(630-647), A#2(649-659), O#3(661-673)]` con la 1.ª oficial (603-613) sin contraparte.
- `[M,M,M]` (agosto-ASE3, Q1-ASE2) → O#1, O#2, O#3 (igual que hoy).

### 1.2 Intent

Corregir el mapeo del interior de `Reporte Componentes R1` con: (a) regla de **canal efectivo** para filas sin D; (b) tipado de seccion **por frontera de cierre**; (c) erratas del T0-Plan45 (§3, §4, §5, §9) sin reescribir historia; (d) retiro de las 2 entradas de `CerosDocumentadosR1` (FP-3); (e) fixtures actualizados CON EVIDENCIA citada (fuente + T0 corregido), nunca por conveniencia; (f) re-corrida completa con evidencia celula a celula en los 5 ASE.

### 1.3 In-scope EXPLICITO

| Frente | Alcance | Criterio de salida |
|---|---|---|
| **H0 - T0 corregido (read-only, cero codigo de produccion)** | Derivar de bytes la matriz corregida por ASE x periodo (julio-2 los 5 ASE + agosto + Q1 read-only) y congelarla en fixture read-only. | Doc T0 corregido + fixture read-only verde; puerta H0→H1. |
| **H1 - Regla canal efectivo + seccion por cierre (contract-first)** | Canal efectivo + tipado por cierre en writer/contrato (sin romper consumidor Plan 35-T2); erratas T0 aplicadas; `CerosDocumentadosR1` sin las 2 filas retiradas (mecanismo conservado). | Casos de aceptacion §S2 verdes + build 0 warnings + greps D-H. |
| **H2 - Fixtures + regresion + re-corrida con evidencia** | Actualizar `ContratoSeccionPeriodFreeH1Tests`, `ContratoSeccionPeriodFreeT0Tests`, `EscritorValoresSlotR1Tests`, `EspejoR1IntegracionTests` con evidencia; `SeccionEspejoR1Tests` (Q1, T1_Ase2Q1) intacto y verde; suite con set-diff; re-corrida end-to-end con verificacion celula a celula + justificacion trazada de todo cambio corrida-8→9. | Set-diff sin rojos nuevos + Q1 verde + SHA intacto + tabla corrida-8→9 justificada. |
| **H3 - Docs** | Erratas en `Docs/T0-Plan45-*.md`, nota de supersesion en `plans/45`, estado en `plans/44`, § Plan 46 en `.opencode/project-context.md`, `Docs/Instructivo-Capa-B.md` si aplica. | Diff de docs citado; cero codigo de produccion en H3. |

### 1.4 Non-goals / out (vinculante)

- Auditoria de la propietaria: sigue PENDIENTE, no se declara paridad.
- Decision Via A/B: solo se proveen numeros frescos post-fix como insumo; no se decide aqui.
- Otras pestanas: fuera.
- Tocar la plantilla oficial: PROHIBIDO (FP-1/FP-3: ni por residuos ni por hueco aparente).
- Re-litigar FP-1/FP-2/FP-3 o los hechos verificados del §0: adoptados.

### 1.5 Riesgos con mitigacion

| ID | Riesgo | Mitigacion |
|---|---|---|
| R-DIVERG | F612 417.519,62→0 cambia los numeros de divergencia (VALIDACION_TOTAL; F612 alimenta F700). | Capturar numeros frescos post-fix como insumo para F2 del abuelo (Via A/B); el aviso de divergencia sigue sin silenciarse. |
| R-FIXTURE | Fixtures que hoy pinchan el comportamiento viejo se actualizan por conveniencia. | Actualizacion SOLO con evidencia citada (fuente + T0 corregido); cada assert cambiado cita fila-fuente + celda-slot + seccion-contrato. |
| R-Q1 | Formas [A,M,A,M] comparten regla con Q1-ASE1/5 → medir impacto. | Neutralidad Q1 (fallback de bloque del Plan 45) preservada; Q1: `T1_Ase2Q1…` verde obligatorio; si algun verde Q1 se moviera, STOP con causa byte-probada. |
| R-CERO | El mecanismo de cero documentado queda sin entradas activas. | Se conserva el mecanismo (por si otra forma lo exige); solo se retiran las 2 filas de FP-3. |
| R-CAMBIO-EXTRA | Cambios de celda corrida-8→9 sin justificacion trazada. | Todo cambio debe justificarse por defecto 1 o 2; cambio inexplicado = STOP. |
| R-FALSO-VERDE | Apagar/sellar/re-etiquetar un gate (criterio E1-1..E1-5 Plan 40). | Prohibido `SeGatea=false`; sub-bloque no evaluable = hallazgo/defecto con frente propio. |

---

## 2. DESIGN

### 2.1 Regla corregida A — canal efectivo (defecto 1)

**Definicion (a implementar):** para una fila SIN D propio, su **canal efectivo** = el D de la fila `Total` que cierra su grupo (la siguiente hacia abajo con D no vacio y ≠ literal `"Total"`); el slot destino debe tener el MISMO canal efectivo. El mismo calculo aplica en slots (oficial) y filas (fuente).

- Filas con D propio (totales, `C=… D=Total`) no cambian.
- La regla vive donde el plan lo exija (writer/contrato: `EncajaEnSlot` / `ConstruirGruposSeccion` / `BuscarGrupo` o predicado puro en Core); D-H intacto (cero strings de periodo/empresa).
- Efecto medido: cada VS/VI/Subsidio cae en su canal (H→H, O→O, T→T, E→E); los totales ya correctos no se mueven.

### 2.2 Regla corregida B — seccion por frontera de cierre (defecto 2)

**Definicion (a implementar):** cada seccion (fuente y oficial) se tipa por la frontera que la CIERRA (la inmediatamente debajo: su fila resumen `Mes`/`Aplicacion…` cuyo importe es su subtotal), no por la que la abre ni por ordinal de tipo.

- ASE5-julio2: la 1.ª (filas 5-13) la cierra Aplicacion (F417.519,62 en su fila 14) → seccion Aplicacion → Aplic#1 oficial (615-628); 603-614 queda en cero.
- ASE1/2/3/4 julio-2 y agosto: su 1.ª cierra con Mes → seccion Oportuno → misma seccion de siempre (identidad, sin cambio).
- La regla se computa desde la forma (geometria pura; cero strings de periodo/empresa — D-H intacta).

**Restriccion de compatibilidad (vinculante):** `SeccionEspejoR1` tiene OTRO consumidor (Plan 35-T2, `R1FirmaInterior.EsAplicacionPorFrontera`, frontera por debajo): NO cambiar la API global `PorFronteraArriba`; introducir el tipado por cierre donde corresponda (writer/contrato) sin romper ese consumidor.

### 2.3 Arquitectura del cambio (piezas y costuras)

- `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs`: `ClasificarSeccionesRaw` (tipado por cierre), `SeccionQueAbre` (→ analogo por cierre), `ClasificarSeccionesContrato`, `EncajaEnSlot` (+ canal efectivo), `ConstruirGruposSeccion` (+ canal efectivo), `BuscarGrupo`, `TryPlanificar`, `LeerSlots`, `Escribir`. Mantener **plan-then-write a nivel de bloque** (neutralidad Q1), values-only, guard anti-formula, decimal exacto, mapeo por encabezado F..Z.
- `Remuneracion.Core/Models/ContratoSeccionPeriodFreeR1.cs`: `Alinear` + `ParticipaDeSeccionConformada` bajo la regla de cierre.
- `Remuneracion.Core/Models/SeccionEspejoR1.cs`: NO romper consumidor Plan 35-T2.
- `Remuneracion.Core/Models/CerosDocumentadosR1.cs`: retirar las 2 entradas de FP-3 (402,40 / 7.697,23); conservar el mecanismo.
- `Remuneracion.Core/Models/R1FirmaInterior.cs`: `EsDatoEmpresa` — baseline Plan 41, intacto.
- Invariantes preservados: values-only; jamas `<f>` fuera de Phase 0 (allow-list 68 celdas); cero `InsertRow`/`DeleteRow`; preservacion shared-formula (Plan 33); D-H; tolerancia ±0.5 solo en comparacion; CRLF sin mezclar; NUNCA `git commit`/`push`; oficial SHA256 intacto; quincena por dominio; tests NO se debilitan; declarada = ultima puerta (Plan 42); prohibido fallback posicional (Plan 41 E1-1) y prohibido sellar/apagar/re-etiquetar (Plan 40 E1-1..E1-5); `<v>` nunca es oraculo.

### 2.4 Fixtures que congelan el mapeo viejo (actualizar CON EVIDENCIA)

- `Remuneracion.IntegrationTests/ContratoSeccionPeriodFreeH1Tests.cs` (p. ej. `AssertSlot("F603", F(5))` en L43; comentarios L42-49/L60-71/L100/L139; test `Contrato_Alinear_MapeosProbados` L157-185 que hoy espera `[1.ª,Aplic#1,…]`).
- `Remuneracion.IntegrationTests/ContratoSeccionPeriodFreeT0Tests.cs` (L71-77: `PrimeraSeccionOficial`, comentario "El tipo de la 1.ª seccion es siempre 'Primera'" → actualizar a la regla de cierre).
- Revisar `EscritorValoresSlotR1Tests.cs` + `EspejoR1IntegracionTests.cs` por asserts que pinen el corrimiento.
- `SeccionEspejoR1Tests.cs` (Q1, `T1_Ase2Q1_NUEVO ESQUEMA`) NO se toca y debe seguir verde.

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda T)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa con trx + **set-diff vs linea base `o43-final-failed.txt` (39 fallos; referencia `p45-h2b.trx` 968/929/39)**: reportar altas/bajas una por una. **Cualquier rojo NUEVO = STOP** y reporte con causa (no se maquilla).
- G-Q1: `T1_Ase2Q1…` verde obligatorio; **si algun verde de Q1 se moviera, STOP con causa byte-probada**.
- G-E2E: corrida end-to-end (regen a temp FRESCA del pipeline sobre `Docs/Prueba Julio-2/Insumos`) + verificacion celula a celula de §S2; **todo cambio de celda entre corrida-8 y corrida-9 debe tener justificacion trazada** (defecto 1 o 2); cambios inexplicados = STOP. ASE2/3/4 (no auditados a fondo) se miden en la re-corrida.
- G-SHA: oficial SHA256 intacto (`3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).
- G-PROC: cero commits; CRLF sin mezclar en lo tocado.
- G-H0: T0 corregido deriva de bytes la matriz por ASE x periodo y la congela en fixture read-only ANTES de tocar produccion (puerta H0→H1).
- G-D-H: cero literales/ramas por periodo o empresa en produccion (grep de cierre).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale).

### S1 — Casos de aceptacion celula a celula (corrida nueva, ASE1/ASE5-julio2)

- S1.1 (ASE1-julio2): F9=0, F10=0, F11=-4.108,55, F12=-4.108,55, F13=0, F14=0, F15=-4.108,55; F6/F7/F8 sin cambio (-7.108,82); F16/F17 sin cambio (-11.217,37). Resto del bloque sin cambio.
- S1.2 (ASE5-julio2, bloque 1.ª): F603..F613=0 y F614=0 (Mes/AFaseo); F619=F621=52.706,16 (H), F620=0; F622=49.921,06, F623=402,40, F624=50.323,46 (O); F625=306.792,77, F626=7.697,23, F627=314.490 (T); F628=F629=417.519,62; F612=0 (antes 417.519,62 — el dinero se ve ahora en F628).
- S1.3 (ASE5-julio2, resto): F630-647 (bloque grande) sin cambio; F648=13.084.233.937,44; F649-659 sin cambio estructural (subsidios con H/O en su canal: F653=-7.906,16, F655=32.235,54, F659=24.329,38); F674=-1.036.320.522,40; cola 675/676 sin cambio.

### S2 — Criterios H1 (implementacion)

- S2.1: filas sin D colocadas por canal efectivo (§2.1) — S1.1/S1.2/S1.3 verdes.
- S2.2: tipado de seccion por cierre (§2.2) — ASE5 603-614 en cero, 615-628 con el bloque; ASE1/2/3/4 identidad.
- S2.3: consumidor Plan 35-T2 intacto (`EsAplicacionPorFrontera`, `PorFronteraArriba` sin cambio de API global; tests `SeccionEspejoR1Tests` verdes).
- S2.4: FP-3 aplicado — texto §9 del T0 retirado (WITHDRAW, no se reporta al dueno) y 2 entradas retiradas de `CerosDocumentadosR1`; mecanismo conservado (vacio/sin esas filas).
- S2.5: `<f>` intacta en todo el diff (grep: cero `new CellFormula` fuera de `EscribirFormulaPreservando`; cero `InsertRow`/`DeleteRow`); Y y Exonera-* intactas; decimal exacto en escritura.

### S3 — Criterios H2 (fixtures + regresion + re-corrida)

- S3.1: fixtures §2.4 actualizados SOLO con evidencia citada (fuente + T0 corregido); cada assert cambiado cita fila-fuente + celda-slot + seccion-contrato.
- S3.2: set-diff vs linea base (39 fallos): altas/bajas una por una; cero rojos nuevos; Q1 verde.
- S3.3: re-corrida end-to-end con tabla corrida-8→corrida-9: cada celda cambiada justificada por defecto 1 o 2; ASE2/3/4 medidos.
- S3.4: numeros frescos de divergencia post-fix capturados como insumo F2 del abuelo (F612→F700); aviso de divergencia sin silenciar.

### S4 — Criterios H3 (docs)

- S4.1: erratas T0 corregidas sin reescribir historia (§3 regla, §4 matriz ASE5, §5 cero documentado, §9 texto retirado).
- S4.2: `plans/45` con nota de supersesion por 46; `plans/44` con estado actualizado.
- S4.3: `.opencode/project-context.md` con § doctrina Plan 46; `Docs/Instructivo-Capa-B.md` actualizado si aplica.

---

## 4. TASKS (fases H0..Hn con puertas)

### H0 — T0 corregido con fixtures read-only (bloqueante; cero codigo de produccion)

1. H0.1 Re-derivar contra bytes (BCL ZIP+XML, sin Excel/COM/OpenXML; helper `dumpxlsx.ps1`): matriz corregida por ASE x periodo (julio-2 los 5 ASE + agosto + Q1 read-only) bajo las reglas §2.1/§2.2; anclas §0.2 re-verificadas; fronteras de cierre citadas por fila.
2. H0.2 Congelar la matriz en fixture read-only (nuevo test T0 del Plan 46; no toca produccion): cada celda con cita fila-fuente + celda-slot + seccion-contrato + canal efectivo.
3. H0.3 Redactar erratas del T0-Plan45 (§3, §4, §5, §9) como texto listo para H3 (sin aplicar aun a docs).
- **Puerta H0→H1:** fixture read-only verde + matriz citada. Sin esto no se abre H1. Si la matriz no absorbe las 3 formas sin ramas, STOP declarado.

### H1 — Implementacion canal efectivo + seccion por cierre (contract-first)

1. H1.1 Core: canal efectivo (predicado puro, sin I/O, sin periodo) + tipado por cierre en `ContratoSeccionPeriodFreeR1` (`Alinear`, `ParticipaDeSeccionConformada`); `SeccionEspejoR1`/`EsAplicacionPorFrontera`/`EsDatoEmpresa` intactos.
2. H1.2 Consumo en `EscritorValoresSlotR1` (`ClasificarSeccionesRaw`, `SeccionQueAbre`→analogo por cierre, `ClasificarSeccionesContrato`, `EncajaEnSlot`, `ConstruirGruposSeccion`, `BuscarGrupo`, `TryPlanificar`); plan-then-write a nivel de bloque (neutralidad Q1) preservado.
3. H1.3 `CerosDocumentadosR1`: retirar las 2 entradas FP-3; conservar el mecanismo.
4. H1.4 Tests focalizados (S1.1/S1.2/S1.3 + Q1 intacto) verdes + build 0 warnings + greps D-H.
- **Puerta H1→H2:** S2.1..S2.5 verdes. Sin esto no se abre H2.

### H2 — Fixtures + regresion + re-corrida con evidencia (calidad senior)

1. H2.1 Actualizar `ContratoSeccionPeriodFreeH1Tests` (L43 `F603`, comentarios L42-49/L60-71/L100/L139, `Contrato_Alinear_MapeosProbados` L157-185) + `ContratoSeccionPeriodFreeT0Tests` (L71-77 `PrimeraSeccionOficial`) + revisar `EscritorValoresSlotR1Tests` / `EspejoR1IntegracionTests` — cada cambio con cita de evidencia.
2. H2.2 Suite completa con trx + set-diff vs `o43-final-failed.txt` (39 fallos): altas/bajas una por una; cero rojos nuevos o STOP con causa; Q1 verde.
3. H2.3 Re-corrida end-to-end (regen a temp FRESCA del pipeline sobre `Docs/Prueba Julio-2/Insumos`) + verificacion celula a celula §S1 en los 5 ASE + tabla corrida-8→9 con justificacion trazada por celda; cambios inexplicados = STOP. Capturar numeros frescos de divergencia (F612→F700) como insumo F2.
4. H2.4 Code-review independiente (spec compliance + calidad): trazabilidad cada cambio → requisito S2/S3; ningun test debilitado (0 `Skip`, 0 `Assert.True(true)`); CRLF sin mezclar; SHA256 intacto.
- **Puerta H2→H3:** review APROBADO (o APROBADO CON OBSERVACIONES cerradas) + set-diff sin rojos nuevos + G-Q1 verde + SHA intacto.

### H3 — Docs (cero codigo de produccion)

1. H3.1 Aplicar erratas en `Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md` (§3 regla, §4 matriz ASE5, §5 cero documentado, §9 texto retirado) como errata fechada, sin reescribir historia.
2. H3.2 Nota de supersesion en `plans/45` (supersedido por 46 en § correpondientes) + estado en `plans/44`.
3. H3.3 `.opencode/project-context.md`: nuevo § doctrina Plan 46 (canal efectivo + seccion por cierre + FP-1/2/3 + formas) + actualizar § Plan 45 donde cite la regla vieja.
4. H3.4 `Docs/Instructivo-Capa-B.md` si aplica (auditoria fila-a-fila por fuente/canal/concepto/seccion-contrato).
- **Puerta H3:** diff de docs citado en el resumen de cierre.

---

## 5. Invariantes vigentes (todos heredados, sin excepcion)

Values-only; jamas `<f>` fuera de Phase 0 (allow-list 68 celdas); cero `InsertRow`/`DeleteRow`; preservacion shared-formula (Plan 33); D-H (cero literales/ramas por periodo o empresa); tolerancia ±0.5 solo en comparacion; CRLF sin mezclar; NUNCA `git commit`/`push`; oficial SHA256 intacto antes/despues; quincena por dominio (`Periodo.NumeroQuincena`), nunca por contenido; tests NO se debilitan; declarada = ultima puerta con doble huella y fail-closed (Plan 42); prohibido fallback posicional (Plan 41 E1-1) y prohibido sellar/apagar/re-etiquetar (Plan 40 E1-1..E1-5); `<v>` nunca es oraculo; Enmienda E-44.1 y puertas del padre/abuelo (auditoria pendiente, Via A/B en F2) intactas y referenciadas, no reabiertas.

Comandos:

```bash
dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx
dotnet test Remuneracion.WinForms/Remuneracion.WinForms.slnx --filter "<focalizados>" --logger "trx"
dotnet test Remuneracion.WinForms/Remuneracion.WinForms.slnx --logger "trx"
```

Rutas explicitas en todo lo nuevo; prohibido `Docs/Insumos` raiz; regens siempre a temp FRESCA con el OFICIAL como unica entrada y la plantilla CERRADA en Excel.

---

## 6. Architecture Validation Certificate

- **SOLID:** SRP (canal efectivo como predicado puro separable / tipado por cierre en contrato / pegado values-only en writer / ceros como DATOS del T0); OCP (nueva forma = fila citada en T0 nuevo, cero ramas nuevas; periodo nuevo = cero cambios por D-H y geometria pura); LSP/DIP (writer depende de `BloqueEspejoAseInputs` + predicados + contrato, no de concretos de periodo); ISP (contratos `IRecaudoReader`/`IPlantillaWriter`/`IDetRetriR10Reader`/`IReporteDivergenciasDeclaradas` intactos; API `PorFronteraArriba` intacta para el consumidor Plan 35-T2).
- **Best Practices:** contract-first (H0 congela bytes corregidos antes de H1); fail-fast accionable (`ERR-SIN-SLOT` residual con ASE+firma+empresa+canal+concepto+seccion+ocurrencia); gates que leen el workbook (texto-`<f>` + literales, nunca `<v>`); casos de aceptacion celula a celula + set-diff contra linea base + G-Q1 + SHA256; CRLF; sin commits.
- **Performance:** canal efectivo resuelto una vez por fila/slot (O(filas + slots), sin re-lectura por celda); tipado por cierre una vez por seccion; agrupacion por (firma, seccion, canal efectivo) con cursor por grupo; lectura acotada en H0 (BCL ZIP+XML); regens a temp fresca sin copiar de mas.

---

## 7. Clarification Gate (que requiere aprobacion del Ingeniero)

1. **Reglas corregidas (§2.1 canal efectivo + §2.2 seccion por cierre):** ratificar ambas reglas y la restriccion de compatibilidad (no romper `PorFronteraArriba`/consumidor Plan 35-T2) antes de abrir H1.
2. **Casos de aceptacion §S1 + justificacion corrida-8→9:** ratificar los valores celula a celula y la regla "cambio inexplicado = STOP".
3. **Erratas T0 + retiro de 2 ceros (FP-3):** ratificar el WITHDRAW del texto §9 y el retiro de las entradas de `CerosDocumentadosR1` (mecanismo conservado).
4. **Puerta H0→H1:** sin H0 verde (matriz citada + fixture read-only) no se implementa. Si H0 declara STOP, no se abre H1.
5. **F2 del abuelo:** Via A/B y auditoria siguen en el 44 con numeros frescos post-hijo; este plan solo entrega numeros, no re-decide. **Q1:** cualquier cambio de comportamiento Q1 mas alla de medir exige autorizacion explicita separada.

Sin estas aprobaciones el plan no avanza de H0 a H1 (1, 4) ni cierra H2 (toda la regresion + 2) ni H3 (3, 5).

---

## 8. Estado

**BORRADOR 2026-10-17** — hijo correctivo del 45 (nieto del 44). Pendiente de aprobacion del Ingeniero. Sin implementacion hasta el Clarification Gate (§7).

---

## 9. CIERRE DE EJECUCION (2026-10-17)

> Ejecutado por el `implementer` tras la aprobación del Ingeniero. **H0/H1/H2/H3 ejecutados**; H2 se
> detuvo en la puerta G-SUITE y el Ingeniero lo resolvió con la **opción A**; todo lo que sigue es
> evidencia verificada (build, trx, corrida, hashes).

### 9.1 H0 — T0 corregido read-only (PUERTA CUMPLIDA)

- Fixture nuevo: `Remuneracion.IntegrationTests/ContratoSeccionPeriodFreeT0Plan46Tests.cs` (BCL ZIP+XML,
  cero producción) → **26/26 verdes**. Deriva contra bytes la matriz por ASE × período (julio-2 5 ASE +
  agosto 5 + Q1 5 read-only) bajo §2.1/§2.2; congela la colocación celda a celda (fila-fuente +
  celda-slot + sección-contrato + canal efectivo) para ASE1-julio2 y ASE5-julio2 y los SIN-SLOT
  documentados. Las 3 formas se absorben sin ramas → **no hay STOP en H0**.

### 9.2 H1 — canal efectivo + sección por cierre (PUERTA CUMPLIDA)

- `Core/Models/ContratoSeccionPeriodFreeR1.cs`: +`SeccionPorCierre` (§2.2) y +`CanalEfectivo` (§2.1);
  `Alinear` intacto.
- `Infrastructure/Excel/EscritorValoresSlotR1.cs`: tipado por cierre + llave
  `(firma + sección + canal)` en `EncajaEnSlot`/`ConstruirGruposSeccion`/`BuscarGrupo`/`TryPlanificar`/
  `LeerSlots`; `SeccionQueAbre` retirado; baseline (neutralidad Q1) intacto.
- `Core/Models/CerosDocumentadosR1.cs`: patrones vacíos (FP-3); mecanismo conservado.
- `SeccionEspejoR1.PorFronteraArriba` y `R1FirmaInterior.EsAplicacionPorFrontera` (Plan 35-T2) INTACTOS.
- Build 0 warnings; tests focalizados verdes; greps de cierre (cero `new CellFormula` fuera del
  preservador, cero `InsertRow`/`DeleteRow`, cero ramas por período/empresa).

### 9.3 H2 — fixtures + regresión + re-corrida (CERRADO con opción A)

**STOP declarado y resuelto**: con el fix, la divergencia real `OCCIDENTE-ASE5` **desapareció**
(`F612 = 0`; `F700 = 304.989.620,81` = medición del administrativo; `VALIDACION_TOTAL!O25/D34/F34`
sellan TRUE con Δ 0; la corrida 9 por CLI cierra con **codigo=0 sin reporte de divergencia**). Eso
dejó 11 rojos nuevos: todos los tests del mecanismo Plan 42/43 que afirmaban "la divergencia existe y
bloquea". El Ingeniero resolvió (2026-10-17) con la **opción A**:

1. **Retiro de la entrada del catálogo** (`DivergenciasDeclaradasCatalogo.Entrada`): se añade el
   MECANISMO MÍNIMO documentado y tested `Retirada` + `FechaRetiro` + `MotivoRetiro` +
   `EvidenciaRetiro` + `RetiroDocumentado` + `Vigente` (+ `Vigentes`/`Retiradas`); el matcher solo
   considera entradas **vigentes**. La entrada `OCCIDENTE-ASE5` queda RETIRADA (fecha 2026-10-17,
   motivo con `F612`→0 y `304.989.620,81`, evidencia con la corrida 9 + §12 del T0-Plan45 + fixtures).
   **El mecanismo de divergencia declarada NO se toca**: solo cambió el DATO.
2. **Rebase de los 11 tests con un FIXTURE SINTÉTICO**
   (`Remuneracion.IntegrationTests/FixtureDivergenciaSintetica.cs`, entrada vigente
   `SINTETICA-C15`, Δ = 1): un lector decorado suma +1 al `Total` del primer `Recaudo *` (celdas
   intactas) → `C15 = (C14 = C9)` FALSE con Δ exacto 1 (y el sellado escribe 0); la clasificación usa
   el MISMO `ValidadorBasico` con un catálogo sintético (sobrecarga pura `Buscar(catalogo, …)`).
   Costuras de test AÑADIDAS (opcionales, defaults = producción): `ComposicionServicios.Crear(validador?, lectorLeaf?)`
   y `EjecutorCli.Ejecutar(..., validador?, lectorLeaf?)` (internal). Cero `Skip`, cero `Assert.True(true)`.
3. **Baja documentada (1)**: `ProcesadorPeriodoTests.Ejecutar_Periodo2026072_ConLectorOracle_VeredictosValidacionesNoVacios`
   pasó de ROJO a VERDE: el test corre Q2 con oráculo y SIN reconocimiento, y pre-fix la corrida
   BLOQUEABA con la declarada (no producía veredictos); con el espejo ya consistente la corrida CIERRA
   y el test verifica los 5 bloques `VALIDACIONES:` como siempre quiso (HU-13/2.7).

**Evidencia final:**
- Suite (`Remuneracion.IntegrationTests/TestResults/p46-cierre.trx`, ver §9.6): **38 Failed / 960 Passed / 998 Total / 0 Skipped**.
- Set-diff vs `o43-final-failed.txt` (39): **0 altas, 1 baja** (la de arriba)
  (`p46-cierre-setdiff.txt`).
- `T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback` → **Passed** (Q1 intocado).
- Build `0 warnings / 0 errores`; oficial SHA256 `3086AC31…C5214` intacto; cero commits; CRLF sin mezclar.
- Re-corrida end-to-end (§S1): CLI 2026072 desde el oficial → **codigo=0**, sin reporte de divergencia;
  §S1.1 (ASE1 `F6..F17`) y §S1.2/§S1.3 (ASE5 `F603..F676`) verificados; sellos `C15/D25/O25/D34/F34 = 1`.

### 9.4 H3 — Docs (APLICADO)

- `Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md` §12 (errata fechada: §3/§4/§5/§9 + canal efectivo + formas).
- `plans/45` nota de supersesión; `plans/44` estado; `.opencode/project-context.md` doctrina Plan 46
  (+ § Plan 45 marcada supersedida); `Docs/Instructivo-Capa-B.md` auditoría R1-interior.

### 9.5 Desviaciones declaradas

- **`SeccionEspejoR1Tests`**: se actualizó UN test (no el Q1)
  `T1_OccidenteAse5_DivergenciaEstructural_Declarada` → `T1_OccidenteAse5_SeccionPorCierre_CoincideConManual`
  (304.989.620,81 ±0,5), porque §S1.2 exige `F612=0`; sin ese cambio habría un rojo nuevo. `T1_Ase2Q1…`
  quedó intacto y verde.
- **Costuras de test** en `ComposicionServicios`/`EjecutorCli`/`Form1`: parámetros OPCIONALES con
  default de producción; no alteran el mecanismo ni el comportamiento productivo.
- **`RegresionEvaluaSellaPlan40Tests.T4` (Q1)**: el caso `2026071` **permanece ROJO** por su causa
  declarada (contrato de sección Q1, `ERR-SIN-SLOT`), fuera del alcance del Plan 46.
- **Errata de trazabilidad del T0-Plan45 (review W-2)**: `ContratoSeccionPeriodFreeT0Tests.cs`
  conservaba la doctrina SUPERSEDIDA del tipado por frontera ARRIBA (comentario de "la 1.ª es siempre
  `Primera`" y expectativas con `Tipo.Primera` en `T0_ReglaPeriodFree_FormasSonSubsecuencias_…`). Se
  añadió una **errata de encabezado** (sin tocar las expectativas byte-level válidas) citando
  `ContratoSeccionPeriodFreeT0Plan46Tests` y `Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md §12`
  (tipado por frontera de CIERRE); el archivo queda declarado como congelamiento HISTÓRICO de la
  aritmética de `Alinear` para las formas viejas.

### 9.6 Review independiente — W-1/W-2 cerrados (2026-10-17)

**W-1 (espejo de la clasificación en el fixture).** `FixtureDivergenciaSintetica` duplicaba línea a
línea `ValidadorBasico.ClasificarDivergenciasDeclaradas` (+ `EvaluarSubBloque` +
`EsErrorDeSubBloqueDeclarado`): si la clasificación real evolucionaba, los tests del mecanismo dejaban
de probar producción en silencio. **Fix:** `ValidadorBasico` gana un constructor con catálogo
OPCIONAL (`null` = `DivergenciasDeclaradasCatalogo.Entradas`, default de producción intacto) que
delega en la sobrecarga pura YA EXISTENTE
`DivergenciasDeclaradasCatalogo.Buscar(catalogo, celda, observado, esperado)`; el fixture se reduce a
`new ValidadorBasico(Catalogo)`. **Espejo eliminado**: la clase decoradora y sus 2 helpers privados
ya no existen; los tests del mecanismo ejercitan la implementación REAL.

**W-2 (trazabilidad de la doctrina supersedida).** Errata de encabezado en
`ContratoSeccionPeriodFreeT0Tests.cs` (ver §9.5): nombra el tipado por CIERRE vigente y el fixture
del Plan 46. Las expectativas byte-level que siguen siendo válidas NO se tocaron.

**Re-verificación (el cambio toca producción: `ValidadorBasico`):**
- Build: **0 warnings / 0 errores**.
- Suite completa (`Remuneracion.IntegrationTests/TestResults/p46-cierre.trx`): **38 Failed / 960
  Passed / 998 Total / 0 Skipped** — MISMO cuadro que `p46-h2b.trx` (conjunto de fallos idéntico,
  comparado uno a uno).
- Set-diff vs `o43-final-failed.txt` (39): **0 altas · 1 baja** (`p46-cierre-setdiff.txt`).
- `T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback` → **Passed** (Q1 intocado).
- SHA oficial `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` intacto; cero
  commits; CRLF sin mezclar; cero `Skip`/`Assert.True(true)`.

**FRENTE 46 CERRADO (2026-10-17).**

