---
titulo: Plan 47 - Revision demas pestanas de pegado con lente R1 (fuente-canal-concepto-componentes)
estado: APROBADO 2026-10-17 (Clarification Gate §7 ratificado) — H0 EJECUTADO (T0 read-only VERDE, 42/42, desvio probado en R2/R4). **H1 EJECUTADO** (canal efectivo R46-1 en R2/R4 + `Otros` + ceros documentados) y **H2 VERIFICADO** (corrida fresca codigo=0; set-diff 38/38 sin altas; `<f>` byte-identica; SHA intacto). **WARNINGS DEL REVIEW CERRADOS 2026-10-17** (W-1 CRLF, W-2 red de seguridad del fail-fast, W-3 canal compuesto por segundo cierre + test anti-orden, S-1 set-diff persistido) con re-verificacion completa; **H3 APLICADO** (doctrina en `.opencode/project-context.md`, `plans/44` follow-up atendido, `Docs/Instructivo-Capa-B.md`). **CERRADO.** Code-review independiente: APROBADO-CON-OBSERVACIONES (observaciones cerradas).
fecha: 2026-10-17
numero: 47
planes_ocupados: 01..46 ocupados; este es el 47
padre: plans/44 - llenado-r1-fuente-canal-concepto-componentes.md (§1.3 follow-up explicito: "Otras pestanas ... misma lente, plan hijo")
madre_doctrina: plans/46 - correccion-mapeo-interior-r1-canal-efectivo-y-seccion-por-cierre.md (CERRADO 2026-10-17; reglas §2.1 canal efectivo + §2.2 seccion por cierre) + plans/45 - contrato-seccion-periodfree-sinc.md (parcialmente supersedido por el 46)
gate_entrada:
  decision_ingeniero_2026-10-17: revisar y —solo si el T0 lo prueba con bytes— corregir el llenado de las demas pestanas de pegado con la misma lente del frente R1 (fuente / canal / concepto / componentes), aplicando las dos reglas nuevas del Plan 46 cuando apliquen
  q1: fuera de alcance (2026071 no se toca; solo se preserva/mide neutralidad)
  plantilla_oficial: NO SE TOCA (SHA256 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` intacto antes/despues)
  suite_base: 38 Failed / 960 Passed / 998 Total (`Remuneracion.IntegrationTests/TestResults/p46-cierre.trx`; referencia de fallos `o43-final-failed.txt`); 0 Skip; tests no se debilitan
  oraculo_final: auditoria de la propietaria (PENDIENTE); prohibido declarar paridad/verde por ausencia; la medicion del administrativo (`Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx`) es solo referencia de desarrollo, jamas oraculo de ejecucion
  sin_commits: NUNCA `git commit`/`push`
---

# Plan 47 - Revision demas pestanas de pegado con lente R1

> SDD completo PROPOSE -> DESIGN -> SPEC -> TASKS. Fuente unica de verdad para la implementacion una vez aprobado ESTE documento. No se implementa sin aprobacion del Ingeniero (punto de bloqueo: sin su OK no se abre H1).
> Lectura obligatoria previa: plan padre completo (44, enmienda E-44.1 y nota F2/Via A-B disuelta) + plan 45 completo (incluida la nota de supersesion por el 46) + plan 46 completo (§0_FP, §2.1/§2.2, §9 cierre con opcion A) + `Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md` (con erratas §12) + `.opencode/project-context.md` (doctrinas Planes 21-46, invariantes, allow-list Phase 0 de 68 celdas) + mapas/escritores vigentes (`Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapDetalleR2R4.cs`, `WorkbookLeafCellMapAjustesSfT.cs`, `WorkbookLeafCellMapReporteBanco.cs`, `WorkbookLeafCellMapBalanceSc.cs`, `WorkbookLeafCellMapDetRetri.cs`, `WorkbookLeafCellMapQ2.cs`, `WorkbookLeafCellMapValidaciones.cs`, `WorkbookLeafCellMap.cs`, `EscritorValoresSlotR1.cs`, `OpenXmlPlantillaWriter.cs` que implementa `IPlantillaWriter`, `IWorkbookLeafWriter`, `IEspejoR1Writer`) + `Docs/Prueba Julio-2/Insumos` (+ subcarpetas por ASE) + corrida corregida de referencia `%TEMP%\opencode\p46-corrida9`.
> Entorno: Windows, repo con espacios en la ruta; la plantilla oficial debe estar CERRADA en Excel al correr tests (FileShare). Oficial: `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`.
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Contexto verificado en bytes (adoptado del pedido; el H0/T0 lo re-verifica por bytes, no se inventa geometria)

### 0.1 Mandato y posicion en la genealogia

- El plan padre (44 §1.3) declara literalmente como follow-up explicito: *"Otras pestanas (R2-detalle, R4-detalle, SALDOS POR NOTA, RETRIBUCION NEGATIVA, AJUSTES-SF-T, BCE, banco, conciliaciones, INTERVENTORIA, AFaseo-consumidores): follow-up explicito (misma lente, plan hijo)"*. Este Plan 47 ES ese plan hijo.
- La lente del frente R1 (cerrado en el 46) es: cada fila fuente `(fuente / canal / concepto / componentes)` coloca **todos sus componentes** en el slot `(firma + seccion + canal)` que le corresponde, por encabezado de columna, values-only, sin insert/delete, sin tocar `<f>`.
- Las dos reglas nuevas del Plan 46 (aplicables "cuando apliquen", segun el pedido aprobado):
  - **R46-1 canal efectivo** (46 §2.1): una fila sin etiqueta propia de canal toma el de la fila `Total` que cierra su grupo.
  - **R46-2 seccion por cierre** (46 §2.2): cada seccion se tipa por la frontera que la **CIERRA** (la inmediatamente debajo), no por la que la abre ni por ordinal de tipo.
- Los dos defectos del 46 que se buscan como gemelos aqui son: **(D1)** filas sin etiqueta que hoy casan **por orden** (corrimiento); **(D2)** bloque mal tipado (frontera que abre vs la que cierra). Mas el precedente metodologico del 46: un "cero documentado" (Plan 45) resulto ser un **falso hueco** al aplicarse la lente nueva — por eso el T0 debe probar el "falso hueco" en cada pestana (ver §1.2-P4).

### 0.2 Inventario duro adoptado (el T0 lo re-verifica; diff read-only oficial → corrida 9 del Plan 46)

Valores que la App cambia por hoja (conteo de celdas con valor distinto entre el oficial y la corrida 9):

| Hoja | Celdas cambiadas | Estado a la entrada |
|---|---|---|
| `Reporte Componentes R1` | 3.496 | Frente **cerrado** en el 46 — solo referencia, no se retoca |
| `Rem. Anticipos R2` | **2.051** | Candidata principal (lente directa) |
| `Reversion Pagos R4` | **828** | Candidata principal (lente directa) |
| `SALDOS POR NOTA` | 237 | Candidata (¿derivada de R2? — lo decide el T0) |
| `REPORTE RECAUDO x BANCO` | 30 | Lente distinta (banco × disposicion) — el T0 la declara dentro/fuera con criterio |
| `BCE SC POR FACT.` | 10 | Agregado/control — el T0 la declara dentro/fuera con criterio |
| `CONSOLIDADO_TOTAL RECAUDO` | 15 | Agregado (arrastre) — fuera salvo arrastre probado |
| `AJUSTES - SF-T` | 6 | Arrastre (D47..D51 = SALDOS + RETRIBUCION) — fuera salvo arrastre probado |
| `VALIDACION_TOTAL` | 8 | Gate/evaluador — fuera (doctrina Plan 40: se evalua, no se pega) |
| `Valida - Control Recaudo` / `Valida -Remunera` | 2 c/u | Control — fuera salvo hallazgo T0 |
| `GERENTES_*` (3 c/u), `REMUNERACION_*` (13 c/u: AD60/N22/N98/N60/N98/AD3) | bajo | Arrastre del periodo — fuera salvo hallazgo T0 |
| `Recaudo Directa Occidente` 36 / `Recaudo ENEL` 36 / `Recaudo EAAB Reciprocidad` 18 / `Recaudo ENERBIT` 8 | bajo | **Aclarar antes de tocar** (¿pegado real? ver §1.2-P5) |
| `DetRetri2026092` 293 / `DetValiRetri2026092` 290 / `Informe AFaseo Recaudo 202609-2` 164 | alto | **Aclarar antes de tocar**: nombres de **septiembre** en el oficial vs corrida **julio-2** (Plan 30 nombres dinamicos; ver §1.2-P5) |
| `INTERVENTORIA`, `RETRIBUCION NEGATIVA`, `ANT EXT-REV`, `ANTICIPOS USUARIOS`, `Recaudo EAAB + Ciud Limp` | 0 (no cambian) | Fuera salvo que el T0 pruebe falso hueco |

**OJO metodologico vinculante (heredado del pedido):** el diff de **formulas es RUIDO por *shared formulas*** → **solo `values_changed` es evidencia**. El T0 que use `<f>` o `<v>` como oraculo queda invalidado (doctrinas Planes 28/29/33/40).

### 0.3 Geometria observada en el oficial (adoptada; el T0 la re-deriva read-only)

- **`Rem. Anticipos R2`:** **C = fuente**, **D = concepto** (`Vlr Servicio` / `Vlr Intereses` / `Total` / `Subsidio(-)/Contribucion(+)` / `Componente`), columnas **E..R = componentes por encabezado** (E Total, F TDF, G TTL, H TVIAT, I Aprovechamiento, J CCSA Prest.Aprov., K Especiales, L TCS, M TLU, N TBL, O TRT, P CCSA No Aprov., Q Deb/Cred, R Deb/Cred-otros). **SIN columna de canal.** Bloques por ASE en `WorkbookLeafCellMapDetalleR2R4.BloquesR2`: ASE1 3-33, ASE2 67-113, ASE3 164-194, ASE4 282-328, ASE5 382-419.
- **`Reversion Pagos R4`:** **A = fuente**, **B = rol** (`Total` / `Componente` / `Subs-Cont`), **C = concepto**, columnas **D..R = componentes por encabezado**. Bloques `BloquesR4`: ASE1 3-15, ASE2 93-111, ASE3 194-206, ASE4 231-249, ASE5 346-358.
- **`SALDOS POR NOTA`:** cuadro **por concesionario** (bloques PROMOAMBIENTAL/LIME/CIUDAD LIMPIA/…), filas `Vlr Servicio` / `Vlr Intereses` / `Componente` / `Subsidio` / `Subs-Cont` / `Total`, columnas **C..O = componentes**; sin fuente/canal explicito (¿derivada de R2? — el T0 lo decide con bytes, no se supone).
- **`REPORTE RECAUDO x BANCO`:** filas = disposicion (1-Aplicados a facturacion, 2-Saldos a favor, 3-Financiaciones, 7-Servicios especiales, Total) por bloques ASE; columnas = **banco** (ENEL, CIUDAD LIMPIA, EAB, NUEVO ESQUEMA, ENERBIT, OCCIDENTE, Total). **Lente distinta** (banco × disposicion): el T0 la declara dentro/fuera con criterio citado (ver §2.1-TABLA).
- **`DetRetri2026092` / `DetValiRetri2026092` / `Informe AFaseo Recaudo 202609-2`:** sellos con codigo de periodo **2026092 (septiembre)** en el oficial; la corrida de referencia es **julio-2 (2026072)**. Rige la doctrina del Plan 30 (`NombresHojaPeriodo`; el runtime NUNCA escribe en hoja de otro periodo; base equivocada = `ERR-PLANTILLA`). El T0 explica estos conteos **antes de proponer tocar nada** (§1.2-P5).

### 0.4 Mapas y escritores vigentes (leerlos; el T0 los contrasta contra bytes)

- `WorkbookLeafCellMapDetalleR2R4.cs` (comenta "Plan 29 T3"; resuelve **fila por LABEL/firma A|B|C|D contra la plantilla**): documenta el riesgo gemelo — *"las filas-fuente sin contraparte se omiten y las filas-destino sin fuente quedan en su 0 de plantilla"* → candidato directo al defecto D1 (corrimiento por orden) y a P4 (falso hueco).
- `WorkbookLeafCellMapAjustesSfT.cs`, `WorkbookLeafCellMapReporteBanco.cs`, `WorkbookLeafCellMapBalanceSc.cs`, `WorkbookLeafCellMapDetRetri.cs`, `WorkbookLeafCellMapQ2.cs`, `WorkbookLeafCellMapValidaciones.cs`, `WorkbookLeafCellMap.cs`, `EscritorValoresSlotR1.cs`, `OpenXmlPlantillaWriter.cs` (implementa `IPlantillaWriter`, `IWorkbookLeafWriter`, `IEspejoR1Writer`).
- Herramientas read-only para el T0 (existen en `%TEMP%\opencode\`; recreables): `dumpxlsx.ps1` (`-Path -Sheet -From -To -Cols -List -OnlyValues`) y `diffxlsx.ps1` (`-A -B -Max -Sheet`). Insumos: `Docs/Prueba Julio-2/Insumos` (+ subcarpetas por ASE). Corrida corregida de referencia: `%TEMP%\opencode\p46-corrida9`.

### 0.5 Linea base y puertas heredadas (no re-litigar; cada gate las re-verifica)

- Linea base de suite a la entrada: **38 Failed / 960 Passed / 998 Total** (`p46-cierre.trx`; referencia de fallos `o43-final-failed.txt`). Criterio vinculante: **set-diff sin altas nuevas**; toda baja se documenta una por una con causa byte-probada; **rojo nuevo = STOP con causa citada** (ver §3-S0).
- Puertas intactas y referenciadas, no reabiertas: auditoria de la propietaria PENDIENTE (no se declara paridad/verde por ausencia); medicion del administrativo solo referencia de desarrollo; Q1 fuera de alcance; plantilla oficial intacta (SHA256 antes/despues); cero commits; D-H (cero literales/ramas por periodo/empresa en produccion); values-only estricto; cero insert/delete; `<f>` intacta (el unico modulo autorizado a tocar formulas es el preservador de *shared formulas*); allow-list Phase 0 de 68 celdas intacta; tolerancia ±0.5 solo en comparacion; CRLF sin mezclar; `<v>` nunca es oraculo; declarada = ultima puerta con doble huella y fail-closed (Plan 42); prohibido fallback posicional (Plan 41 E1-1) y prohibido sellar/apagar/re-etiquetar gates (Plan 40 E1-1..E1-5).

---

## 1. PROPOSE

### 1.1 Intent

Revisar —y **solo si el T0 lo prueba con bytes**— el llenado de **las demas pestanas de pegado** con la misma lente del frente R1 (fuente / canal / concepto / componentes), aplicando R46-1 (canal efectivo) y R46-2 (seccion por cierre) **cuando apliquen**. Cada pestana termina en uno de dos estados, sin tercero: **(a)** corregida por regla geometrica con evidencia T0, o **(b)** declarada **sin hallazgo con evidencia** (no se toca). Prohibido el tercer estado "se toco por intuicion".

### 1.2 In-scope EXPLICITO (lo que el plan DEBE resolver; prohibido suponer)

| ID | Pregunta vinculante | Criterio de salida |
|---|---|---|
| **P1 — alcance de la lente** | ¿Que pestanas **realmente** tienen la lente (fuente/canal/concepto/componentes) y cuales no? | Tabla §2.1 con veredicto por pestana + **criterio citado** para las que quedan fuera (CONSOLIDADO, AJUSTES-SF-T, VALIDACION_*, GERENTES_*, REMUNERACION_*, BCE/banco, DetRetri: agregados/control/otra lente — ver §2.1). Sin criterio citado no hay "fuera". |
| **P2 — corrimiento por orden** | ¿Hay en R2/R4/SALDOS POR NOTA **filas-fuente sin etiqueta que hoy casan por orden** (corrimiento del mismo tipo que el defecto D1 del 46)? | Contraste fila a fila contra el insumo real + corrida 9; cada corrimiento probado cita fila-fuente + celda-slot + regla que lo tapa; si no hay, se declara sin hallazgo con evidencia. |
| **P3 — bloque mal tipado** | ¿Hay **bloques mal tipados** (frontera que abre vs la que cierra, defecto D2 del 46)? | Cada bloque tipado por su frontera de cierre citada; si la tipificacion vigente ya es por cierre, se declara sin hallazgo con evidencia. |
| **P4 — falso hueco** | ¿Hay **filas-destino sin fuente que en realidad SI tienen contraparte** (el "falso hueco" del 46)? | Por cada "omitida/sin contraparte" del mapa vigente: o se prueba la ausencia con bytes (sigue omitida) o se prueba la contraparte (se corrige). Prohibido mantener un cero documentado sin re-probarlo bajo la lente nueva. |
| **P5 — explicar antes de tocar** | `Recaudo *` (¿pegado real?) + `DetRetri2026092`/`DetValiRetri2026092`/`Informe AFaseo Recaudo 202609-2` (nombres de **septiembre** en el oficial vs corrida **julio-2**; Plan 30 nombres dinamicos) | Explicacion citada **antes** de cualquier propuesta de cambio: que escribe el runtime en cada una, desde que insumo, y por que el diff muestra esos conteos. Si la explicacion es "otra quincena/periodo/base", se declara fuera con criterio y no se toca. |
| **P6 — matrices** | Matriz por pestana `(etiquetas de fila) → (slot de fila)` y `(encabezado de componente) → (columna destino)`, con **firma/clave** de cada fila y regla de desempate cuando falte etiqueta | Matriz congelada en fixture read-only (§2.2, §4-H0); sin matriz verde no hay H1. |

### 1.3 Non-goals / out (vinculante)

- **Q1 (2026071): fuera de alcance.** No se cambia comportamiento Q1; Q1 solo se preserva/mide (neutralidad). Cualquier cambio de comportamiento Q1 exige autorizacion explicita separada.
- **Plantilla oficial: NO SE TOCA** (ni por residuos ni por hueco aparente). SHA256 intacto antes/despues.
- **Auditoria de la propietaria: PENDIENTE.** No se declara paridad/verde por ausencia. La medicion del administrativo es solo referencia de desarrollo.
- **Sin commits.** NUNCA `git commit`/`push`.
- **R1 interior: cerrado.** `Reporte Componentes R1` (3.496) es solo referencia/oraculo de arrastre; no se retoca.
- **Gates/evaluadores y agregados no se "pegan":** `VALIDACION_TOTAL` se evalua (Plan 40), no se pega; `CONSOLIDADO_TOTAL RECAUDO` / `AJUSTES - SF-T` / `REMUNERACION_*` / `GERENTES_*` son arrastre/agregado — solo se tocan si el T0 prueba arrastre roto con bytes (y entonces el fix vive en el origen, no en el agregado).

### 1.4 Riesgos con mitigacion

| ID | Riesgo | Mitigacion |
|---|---|---|
| R-ORDEN | El mapa vigente "omite sin contraparte / deja en 0" esconde un corrimiento por orden (gemelo D1). | T0 fila a fila fuente↔slot con firma/clave citada; cada omitida re-probada (P4); corrimiento probado = fix por regla geometrica, no por parche. |
| R-SECCION | Bloque tipado por la frontera que abre en vez de la que cierra (gemelo D2). | Re-tipado por cierre citado por bloque (P3); identidad cuando ya es por cierre (sin hallazgo con evidencia). |
| R-FALSO-HUECO | Mantener un cero/omision que la lente nueva revela como contraparte real. | P4 bloqueante: ninguna omision sobrevive sin prueba de ausencia en bytes. |
| R-OTRA-LENTE | Aplicar la lente R1 a una pestana de otra lente (banco × disposicion; agregados; controles). | P1 con criterio citado por pestana (§2.1); fuera = fuera con criterio, no por intuicion. |
| R-PERIODO | Tocar `DetRetri2026092`/septiembre corriendo julio-2 (Plan 30). | P5 primero: explicar con bytes que escribe el runtime en cada hoja y desde que insumo; base/periodo equivocado = `ERR-PLANTILLA`, nunca escritura cruzada. |
| R-ATRIBUCION | Atribuir a colocacion lo que es plantilla (o viceversa). | T0 aisla por bytes: grep de consumidores `<f>` (maestros shared incluidos) + traza lector (filas emitidas) vs escritor (slots consumidos) por (firma, seccion, ocurrencia); veredicto citado por celda; sin veredicto no hay fix. |
| R-FALSO-VERDE | Apagar/sellar/re-etiquetar un gate para pasar (E1-1..E1-5 Plan 40). | Prohibido `SeGatea=false`; sub-bloque no evaluable = hallazgo/defecto con frente propio, nunca verde. |
| R-Q1 | Contagio a Q1 al implementar. | Q1 fuera de alcance + gate G-Q1 (§3-S0); cualquier movimiento Q1 = STOP con causa. |
| R-FIXTURE | Fixtures que pinchan el comportamiento viejo actualizados por conveniencia. | Actualizacion SOLO con evidencia citada (fila-fuente + celda-slot + seccion/firma + canal efectivo); cada assert cambiado cita bytes. |
| R-CAMBIO-EXTRA | Cambios de celda corrida-9→nueva sin justificacion trazada. | Todo cambio justificado por P2/P3/P4; cambio inexplicado = STOP. |

---

## 2. DESIGN

### 2.1 Tabla de alcance por pestana (el T0 la confirma o corrige con bytes; sin T0 verde no hay H1)

| Pestana | Lente a la entrada | Veredicto esperado (a confirmar por T0) | Criterio |
|---|---|---|---|
| `Rem. Anticipos R2` | **Directa**: filas (C=fuente, D=concepto, firma A|B|C|D) × columnas (E..R por encabezado) | **DENTRO** (candidata principal P2/P3/P4) | Pegado fila-a-fila por label + encabezado; sin canal explicito (¿aplica R46-1 por grupo? lo decide el T0) |
| `Reversion Pagos R4` | **Directa**: filas (A=fuente, B=rol, C=concepto, firma A|B|C) × columnas (D..R por encabezado) | **DENTRO** (candidata principal P2/P3/P4) | Idem R2; bloques `BloquesR4` citados |
| `SALDOS POR NOTA` | **Parcial**: cuadro por concesionario, filas concepto, columnas C..O componentes; sin fuente/canal explicito | **DENTRO si el T0 prueba derivacion** (¿de R2?) con bytes; si no, fuera con criterio | No suponer derivacion: el T0 cita insumo→celda o declara otra lente |
| `REPORTE RECAUDO x BANCO` | **Otra lente**: banco × disposicion (filas disposicion por bloques ASE; columnas banco) | **FUERA** salvo hallazgo T0 (P1 con criterio) | Lente distinta a fuente/canal/concepto/componentes |
| `BCE SC POR FACT.` | Agregado/control (10 celdas) | **FUERA** salvo hallazgo T0 | Agregado; el fix viviria en el origen, no aqui |
| `CONSOLIDADO_TOTAL RECAUDO` (15) / `AJUSTES - SF-T` (6) | Arrastre (D104.. = suma; D85.. = SALDOS + RETRIBUCION) | **FUERA** salvo arrastre roto probado | Agregados; no se pegan |
| `VALIDACION_TOTAL` (8), `Valida -*` (2 c/u) | Gate/evaluador + control | **FUERA** | Se evalua (Plan 40), no se pega |
| `GERENTES_*`, `REMUNERACION_*` | Arrastre del periodo / metadatos | **FUERA** salvo hallazgo T0 | Arrastre; periodo por dominio |
| `Recaudo *` (36/36/18/8) | Por aclarar (P5) | **Explicar antes de tocar** | ¿Pegado real? ¿que insumo? ¿que columnas por quincena? |
| `DetRetri2026092` / `DetValiRetri2026092` / `Informe AFaseo Recaudo 202609-2` | Por aclarar (P5; nombres septiembre vs corrida julio-2) | **Explicar antes de tocar** (Plan 30) | Nombres dinamicos por periodo; prohibida escritura cruzada |
| `INTERVENTORIA`, `RETRIBUCION NEGATIVA`, `ANT EXT-REV`, `ANTICIPOS USUARIOS`, `Recaudo EAAB + Ciud Limp` (0 cambios) | Intactas | **FUERA** salvo falso hueco probado (P4) | Cero cambios ≠ cero prueba: el T0 las barre contra el insumo |

### 2.2 Reglas a aplicar (cuando apliquen; el T0 decide el "cuando")

- **R46-1 canal efectivo (analogia):** donde una fila-detalle carezca de etiqueta propia que la ubique dentro de su grupo (en R1 era D; en R2/R4/SALDOS el T0 identifica la columna-etiqueta analoga si existe), su ubicacion = la de la fila `Total` que cierra su grupo. El T0 cita la columna-etiqueta por hoja o declara que no aplica (y por que).
- **R46-2 seccion por cierre (analogia):** donde existan bloques con frontera-resumen (en R1 era `Mes`/`Aplicacion`; en R2/R4/SALDOS el T0 identifica la fila-frontera analoga si existe), cada bloque se tipa por la frontera que lo **CIERRA**. El T0 cita la frontera por bloque o declara que no aplica (y por que).
- **Firma/clave por hoja (P6):** el T0 congela por pestana DENTRO la clave efectiva — p. ej. R2 `(C-fuente, D-concepto, firma-A|B|C|D, seccion/bloque-ASE)`; R4 `(A-fuente, B-rol, C-concepto, firma-A|B|C, seccion/bloque-ASE)` — con regla de desempate citada cuando falte etiqueta (¿canal/grupo efectivo? ¿omision probada? ¿`ERR-SIN-SLOT` analogo?). **Prohibido casar por orden fisico** (Plan 41 E1-1): sin clave no hay colocacion; fila sin slot probado = omision probada o fail-fast accionable que nombra hoja+ASE+fila (el H1 tipifica el error sin inventar codigos fuera del vocabulario vigente).
- **Columnas por encabezado, nunca por indice:** R2 E..R y R4 D..R ya se resuelven por encabezado en el mapa vigente; el T0 re-verifica cada par (encabezado-fuente → columna-destino) contra bytes y lista los sin-contraparte (¿0 explicito legitimo? ¿falso hueco?).

### 2.3 Arquitectura del cambio (solo si el T0 prueba desvio; contract-first)

- **Core (puro, sin I/O, sin periodo/empresa — D-H):** predicados puros analogos a `ContratoSeccionPeriodFreeR1.SeccionPorCierre` / `.CanalEfectivo` donde el T0 pruebe que aplican (p. ej. `...R2R4` si P2/P3 lo exigen). Reutilizar los contratos puros existentes donde apliquen; nada de parches por periodo/empresa.
- **Infrastructure (values-only):** consumo en el writer/lector que el T0 senale (mapas `WorkbookLeafCellMapDetalleR2R4` / `AjustesSfT` / resto + `OpenXmlPlantillaWriter`), preservando guard anti-formula (`celda-destino formula ⇒ ERR-PLANTILLA`), decimal exacto, cero `InsertRow`/`DeleteRow`, preservacion de *shared formulas* (Plan 33), allow-list Phase 0 intacta.
- **Compatibilidad:** `SeccionEspejoR1.PorFronteraArriba`, `R1FirmaInterior.EsAplicacionPorFrontera` / `EsDatoEmpresa`, `SinonimosEmpresaR1`, consumidor Plan 35-T2 y todo el frente R1 del 46 quedan **intactos** (este plan no reabre R1).
- **Neutralidad Q1:** Q1 fuera de alcance; el H1 preserva plan-then-write / comportamiento historico donde exista (gemelo de la neutralidad Q1 del 45/46) y el gate G-Q1 lo verifica.

### 2.4 Fixtures y evidencia que el T0 debe congelar

- Fixture read-only nuevo (BCL ZIP+XML, cero produccion; gemelo de `ContratoSeccionPeriodFreeT0Plan46Tests`, 26 verdes): matriz por pestana DENTRO × ASE con cada celda citada (fila-fuente + celda-slot + firma/clave + canal/grupo efectivo + frontera de cierre) + lista de omitidas probadas + lista de fuera-con-criterio (P1) + explicacion P5.
- Referencias de conteo a re-verificar: R2 2.051 · R4 828 · SALDOS 237 · banco 30 · BCE 10 · resto §0.2.

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda T)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa con trx + **set-diff vs linea base `p46-cierre.trx` (38/960/998; referencia `o43-final-failed.txt`)**: altas/bajas una por una. **Rojo nuevo = STOP con causa citada** (no se maquilla). 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados.
- G-Q1: Q1 intacto (fuera de alcance); **cualquier movimiento Q1 = STOP con causa byte-probada**.
- G-E2E: corrida fresca a temp (regen del pipeline sobre `Docs/Prueba Julio-2/Insumos` con el OFICIAL como unica entrada) + verificacion fila a fila §S1; **todo cambio corrida-9→nueva justificado por P2/P3/P4**; cambio inexplicado = STOP.
- G-SHA: oficial SHA256 intacto (`3086AC31…C5214`) antes/despues.
- G-PROC: cero commits; CRLF sin mezclar en lo tocado.
- G-H0: T0 read-only deriva la matriz §2.1/§2.2 y la congela en fixture read-only ANTES de tocar produccion (puerta H0→H1). **Sin T0 verde no hay H1. Si el T0 NO encuentra desvio en una pestana, se declara sin hallazgo con evidencia y no se toca.**
- G-D-H: cero literales/ramas por periodo o empresa en produccion (grep de cierre).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale; diff de `<f>` es ruido por shared — solo `values_changed` es evidencia).
- G-FORMULA: `<f>` byte-identica en todo el diff salvo allow-list Phase 0 (grep: cero `new CellFormula` fuera del preservador; cero `InsertRow`/`DeleteRow`).

### S1 — Criterios H0/T0 (read-only, cero codigo de produccion)

- S1.1: matriz por pestana (P6) congelada en fixture read-only con firma/clave por fila y regla de desempate citada.
- S1.2: veredicto P2 por pestana DENTRO (¿filas sin etiqueta casando por orden? citar cada una o declarar ausencia con evidencia).
- S1.3: veredicto P3 por bloque (¿mal tipado abre-vs-cierra? citar cada uno o declarar identidad con evidencia).
- S1.4: veredicto P4 por omitida (cada "sin contraparte / queda en 0" del mapa vigente: ausencia probada o contraparte encontrada).
- S1.5: tabla P1 (dentro/fuera por pestana con criterio citado) + explicacion P5 (`Recaudo *`, `DetRetri2026092`/septiembre) **antes** de cualquier propuesta de cambio.
- S1.6: medicion fila a fila contra insumo y contra la corrida 9; veredicto por celda.

### S2 — Criterios H1 (correccion minima por regla geometrica, solo donde el T0 probo desvio)

- S2.1: cada fix cita T0 (fila-fuente + celda-slot + firma/clave + canal/grupo efectivo + frontera de cierre); reutiliza contratos puros de Core; cero parches por periodo/empresa.
- S2.2: pestanas sin hallazgo **intactas** (cero cambios de produccion en ellas; el fixture T0 es su prueba).
- S2.3: frente R1 del 46 intacto; consumidor Plan 35-T2 intacto; `DivergenciasDeclaradasCatalogo` intacto (aqui no se decide nada de declaradas).
- S2.4: `<f>` intacta; Y/Exonera-* y columnas sin contraparte intactas salvo falso hueco probado; decimal exacto; guard anti-formula intacto.
- S2.5: build 0 warnings + greps D-H + G-FORMULA.

### S3 — Criterios H2 (re-verificacion)

- S3.1: corrida fresca a temp + matriz fila a fila verde + `<f>` byte-identica + set-diff vs 38/960/998 + Q1 intacto + SHA intacto.
- S3.2: tabla corrida-9→nueva con justificacion trazada por celda (P2/P3/P4); cambio inexplicado = STOP.
- S3.3: numeros frescos capturados como insumo (sin re-decidir nada fuera de este plan).

### S4 — Criterios H3 (docs/cierre)

- S4.1: doctrina del plan en `.opencode/project-context.md` (solo lo probado).
- S4.2: estados en `plans/44` (follow-up atendido) sin reescribir historia; desviaciones declaradas con fecha/motivo/evidencia.
- S4.3: diff de docs citado; cero codigo de produccion en H3.

---

## 4. TASKS (fases H0..H3 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero codigo de produccion)

1. H0.1 Re-verificar por bytes (BCL ZIP+XML; `dumpxlsx.ps1` / `diffxlsx.ps1`): inventario §0.2 (R2 2.051 · R4 828 · resto) oficial→corrida 9; geometria §0.3 por hoja; que escribe el runtime en cada hoja P5 y desde que insumo.
2. H0.2 Matriz P6 por pestana candidata: `(etiquetas de fila) → (slot de fila)` + `(encabezado de componente) → (columna destino)` con firma/clave y desempate citado; veredictos P1/P2/P3/P4 por pestana/bloque/fila.
3. H0.3 Congelar en fixture read-only nuevo (cero produccion) + doc T0 con tablas y veredicto por celda; medicion fila a fila vs insumo y vs corrida 9.
- **Puerta H0→H1:** fixture read-only verde + matriz citada + veredictos P1..P5. Sin esto no se abre H1. Pestana sin hallazgo = no se toca. Si el T0 no absorbe una forma sin ramas, STOP declarado.

### H1/F1 — Correccion minima (solo donde el T0 probo desvio)

1. H1.1 Core: predicados puros analogos a R46-1/R46-2 donde el T0 pruebe que aplican (D-H; reutilizar contratos puros).
2. H1.2 Consumo values-only en el writer/mapa senalado (guard anti-formula, decimal exacto, encabezado-no-indice, cero insert/delete).
3. H1.3 Tests focalizados por pestana corregida (verdes) + build 0 warnings + greps D-H + G-FORMULA.
- **Puerta H1→H2:** S2.1..S2.5 verdes. Sin esto no se abre H2.

### H2/F2 — Re-verificacion

1. H2.1 Corrida fresca a temp + matriz fila a fila verde + `<f>` byte-identica + set-diff vs 38/960/998 + Q1 intacto + SHA intacto.
2. H2.2 Tabla corrida-9→nueva justificada por celda; cambios inexplicados = STOP.
3. H2.3 Code-review independiente (spec compliance + calidad): trazabilidad cada cambio → P2/P3/P4; ningun test debilitado; CRLF; SHA.
- **Puerta H2→H3:** review APROBADO (o APROBADO CON OBSERVACIONES cerradas) + set-diff sin rojos nuevos + G-Q1 + SHA.

### H3/F3 — Docs/cierre

1. H3.1 Doctrina probada en `.opencode/project-context.md`; estados en `plans/44`; desviaciones declaradas con fecha/motivo/evidencia.
2. H3.2 `Docs/Instructivo-Capa-B.md` si aplica (auditoria fila a fila por las claves nuevas).
- **Puerta H3:** diff de docs citado en el resumen de cierre.

---

## 5. Riesgos (tabla consolidada + gates por fase)

Ver §1.4 (R-ORDEN, R-SECCION, R-FALSO-HUECO, R-OTRA-LENTE, R-PERIODO, R-ATRIBUCION, R-FALSO-VERDE, R-Q1, R-FIXTURE, R-CAMBIO-EXTRA). Gates por fase:

- **Gate H0→H1:** fixture T0 verde + matriz citada + P1..P5 resueltos. Sin T0 verde no hay H1.
- **Gate H1→H2:** S2 verde (solo pestanas con desvio probado) + build 0w + greps D-H + G-FORMULA.
- **Gate H2→H3:** set-diff sin altas + Q1 intacto + SHA intacto + tabla corrida-9→nueva justificada + review aprobado.
- **Gate H3:** diff de docs citado; cero codigo de produccion en H3.

---

## 6. Architecture Validation Certificate

- **SOLID:** SRP (T0 read-only / predicados puros en Core / pegado values-only en Infrastructure / matrices como DATOS del T0); OCP (pestana o forma nueva = fila citada en T0 nuevo, cero ramas nuevas; periodo nuevo = cero cambios por D-H y geometria pura); LSP/DIP (writers dependen de mapas + predicados + `IPlantillaWriter`/`IWorkbookLeafWriter`, no de concretos de periodo); ISP (contratos `IRecaudoReader`/`IPlantillaWriter`/`IWorkbookLeafWriter`/`IEspejoR1Writer`/`IDetRetriR10Reader` intactos).
- **Best Practices:** contract-first (H0 congela bytes antes de H1); fail-fast accionable (hoja+ASE+fila/firma); gates que leen el workbook (texto-`<f>` + literales, nunca `<v>`); aceptacion fila a fila + set-diff contra 38/960/998 + G-Q1 + SHA256; CRLF; sin commits.
- **Performance:** etiquetado/clave una vez por fila/slot (O(filas + slots), sin re-lectura por celda); agrupacion por (firma, seccion/bloque, canal/grupo) con cursor por grupo donde aplique; lectura acotada en H0 (BCL ZIP+XML); regens a temp fresca.

---

## 7. Clarification Gate (requiere aprobacion del Ingeniero; sin esto no se abre H1)

1. **Alcance P1 + P5:** ratificar la tabla §2.1 (que pestanas entran/salen y con que criterio) y la explicacion de `Recaudo *` + hojas `2026092`/septiembre antes de tocar nada.
2. **Reglas analogas R46-1/R46-2:** ratificar que aplican donde el T0 pruebe desvio, y que pestana sin hallazgo no se toca.
3. **Puerta H0→H1:** sin T0 verde (matriz citada + fixture read-only + P1..P5) no se implementa. Si H0 declara STOP, no se abre H1.
4. **Q1 / plantilla / auditoria / commits:** ratificar que Q1 sigue fuera, el oficial intacto, la auditoria pendiente (sin paridad por ausencia) y cero commits.
5. **Linea base de suite:** ratificar 38/960/998 como base del set-diff y que rojo nuevo = STOP con causa citada.

Sin estas aprobaciones el plan no avanza de H0 a H1 (1, 3, 4, 5) ni cierra H2 (toda la regresion + §S3) ni H3 (4, 5).

---

## 8. Estado

**BORRADOR 2026-10-17** — plan hijo del 44 (misma lente, demas pestanas). Pendiente de aprobacion del Ingeniero. Sin implementacion hasta el Clarification Gate (§7).

---

## 9. EJECUCION H0 — T0 read-only (2026-10-17)

> Ejecutado por el `implementer` tras la aprobacion del Ingeniero (Clarification Gate §7 ratificado).
> **H0 completo; H1 NO abierto** (requiere autorizacion). Evidencia detallada, tablas y citas por celda
> en `Docs/T0-Plan47-Revision-Pestanas-LenteR1.md`.

### 9.1 Entregables

- **Fixture read-only nuevo:** `Remuneracion.IntegrationTests/RevisionPestanasLenteR1T0Plan47Tests.cs`
  (BCL ZIP+XML; cero writers/readers de produccion en la derivacion) → **42/42 verdes**. Congela: la
  matriz P6 por ASE (firma `A|B|C|D` en R2 / `A|B|C` en R4, pares LCS, filas sin fuente, omitidas con
  valor), los encabezados fuente↔destino, el canal efectivo (R46-1) de la fila-fuente vs el del slot, el
  tipado por cierre (P3), las omitidas/falsos huecos (P4), `Recaudo *` como pegado real desde
  `Conciliaciones/RESUMEN MES` (P5), el placeholder de septiembre de las 3 hojas de periodo (P5) y el
  impacto (ninguna `<f>` consume filas sin etiqueta).
- **Doc T0:** `Docs/T0-Plan47-Revision-Pestanas-LenteR1.md` (inventario oficial→corrida 9, P1..P5,
  medicion fila a fila vs insumo y vs corrida 9, gates y veredicto).

### 9.2 Veredicto de la puerta H0→H1

**(a) DESVIO PROBADO (candidato a H1):**

1. **`Rem. Anticipos R2`** — corrimiento por orden (gemelo D1) en **ASE1, ASE2, ASE4 y ASE5**: el
   `Vlr Servicio`/`Vlr Intereses` del canal **E** cae en las filas del grupo **A/EAB** (`F3/F4`,
   `F67/F68`, `F282/F283`, `F382/F383`) y las filas propias del canal E (`E7/E8`, `E71/E72`,
   `E286/E287`, `E386/E387`) quedan en el 0 de plantilla = **falso hueco** con contraparte real.
   Ademas **omisiones silenciosas**: ASE2 (3 filas: `Vlr Intereses` canal 2 = 406,42; grupo canal 9 =
   831.919,29) y ASE4 (6 filas: canal 2 = −8.785,59; canal 9 = 20.150.395,25 y 5.663.564,75).
2. **`Reversion Pagos R4`** — corrimiento por orden en **ASE2** (el grupo **OCCIDENTE** de la fuente
   cae en `D99:E103`, filas del grupo **NUEVO ESQUEMA**; `D105:D109` = 0) y **falso hueco de columna**
   en **ASE4** (la fuente trae `Otros` con valores y el mapa no la declara; `R232` es VALOR sin `<f>`).
3. La clave vigente casa **por orden fisico** en filas de firma duplicada (Plan 41 E1-1): la correccion
   debe usar `(canal efectivo R46-1 + firma)` y decidir el tratamiento de las filas sin slot (reporte
   documentado vs fail-fast accionable) — **no se toca nada en H0**.

**(b) SIN HALLAZGO CON EVIDENCIA (no se tocan):** `SALDOS POR NOTA` (mapeo por titulo + `Vlr Intereses`
opcional; `Especiales` ausente probado), `REPORTE RECAUDO x BANCO` (otra lente; hoy alineada por
etiqueta ASE, riesgo posicional documentado), `BCE SC POR FACT.`, `CONSOLIDADO_TOTAL RECAUDO`,
`AJUSTES - SF-T`, `VALIDACION_*`, `Valida -*`, `GERENTES_*`, `REMUNERACION_*`, `INTERVENTORIA`,
`RETRIBUCION NEGATIVA`, `ANT EXT-REV`, `ANTICIPOS USUARIOS`, `Recaudo *` y las 3 hojas de periodo
(renombradas/reescritas por el preparador de base; cero escritura cruzada).

**P3 = identidad**: los 10 bloques de R2/R4 ya se tipan por la frontera que cierra (`A=Total`).
**No hay STOP en H0**: la matriz se absorbe sin ramas por periodo/empresa.

**Alcance propuesto para H1 (minimo, requiere autorizacion):** solo R2 y R4; predicado puro de canal
efectivo en Core (analogo a `ContratoSeccionPeriodFreeR1.CanalEfectivo`, D-H); conservar intactos los
totales y todo el frente R1 del 46; decidir con el Ingeniero (i) el tratamiento de las filas sin slot
(ASE2/ASE4) y (ii) si `Otros` de R4-ASE4 se escribe en `R` o se declara ausente formalmente.

### 9.3 Gates

- **G-BUILD:** 0 warnings / 0 errores.
- **G-SUITE:** `Remuneracion.IntegrationTests/TestResults/p47-h0.trx` → **38 Failed / 1002 Passed /
  1040 Total / 0 Skipped**; set-diff vs `p46-cierre.trx` → **0 altas / 0 bajas**
  (`p47-h0-setdiff.txt`); los 42 nuevos entran en Passed (960 + 42).
- **G-Q1:** `T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback` → Passed.
- **G-SHA:** oficial `3086AC31…C5214` intacto antes/despues.
- **G-PROC:** cero commits; CRLF sin mezclar; cero codigo de produccion tocado (H0 read-only).
- **Q1 / plantilla / auditoria / declaradas:** intactas y referenciadas, no reabiertas.

### 9.4 Hallazgos de herramienta (no de producto)

- Los conteos del §0.2 para `DetRetri2026092` (293) / `DetValiRetri2026092` (290) / `Informe AFaseo
  Recaudo 202609-2` (164) son un **artefacto** de la version previa de `diffxlsx.ps1` (hoja ausente en
  B → comparaba contra el `$vb` de la hoja anterior). La comparacion correcta (par renombrado) da
  **19 / 8 / 0**. El T0 usa `%TEMP%\opencode\p47-diff.ps1` (clasifica solo_A/solo_B) y deja el ruido
  documentado.
- El diff de **formulas** sigue siendo RUIDO por *shared formulas* + los 6 tokens `<f>` que reescribe el
  preparador de base (`BCE!H3..`, `VALIDACION_*!C4..`): solo `values_changed` es evidencia.

---

## 10. EJECUCION H1 + H2 — correccion R2/R4 con canal efectivo (2026-10-17)

> Ejecutado por el `implementer` con la **autorizacion H1 completa en R2/R4** del Ingeniero (puntos
> 1–4). Detalle por celda, trazas `archivo:linea`, tabla corrida-9→nueva y gates en
> `Docs/T0-Plan47-Revision-Pestanas-LenteR1.md` §12.

### 10.1 Que cambio (contract-first)

- **Core (puro, D-H)** — `CanalEfectivoDetalleR2R4` (nuevo): `Rotulo` (A–D), `EtiquetaDeGrupo`,
  `TieneEtiquetaPropia` y `CanalEfectivo` (R46-1 analogo: la etiqueta de grupo de la fila `Total` que
  **CIERRA** el grupo, tanto en la fuente como en el oficial).
- **Core (puro)** — `CerosDocumentadosDetalleR2R4` (nuevo, doctrina E-44.1): 11 claves / **14 filas**
  (9 de julio-2 + 5 de agosto-2) con motivo y evidencia citada; `CantidadDocumentada` acota el consumo.
- **Infrastructure (values-only)** — `WorkbookLeafCellMapDetalleR2R4`: `ColumnasEtiquetaPropia/Grupo`
  por hoja + `("Otros","R")` en `ColumnasR4` (13 columnas; `SERVICIO ESPECIALES P` y
  `Ajuste Recargo Aseo Q` quedan fuera por ausencia probada).
- **Infrastructure** — `OpenXmlPlantillaWriter`: `EmparejarBloque` ahora casa por
  **(firma + canal efectivo)**; se RETIRARON `AlinearFirmas`, `LeerFirmaDestino` y `NormalizarFirma`;
  sin slot ⇒ cero documentado de la lista o `ERR-SIN-SLOT` accionable (hoja + ASE + fila + firma +
  canal). Cero `<f>`, cero insert/delete, guard anti-formula intacto.
- **Fixtures (R-FIXTURE, con evidencia citada)** — `DetalleR2R4EscrituraTests` (3 tests y 2 helpers
  actualizados: `E7`/`F7`/`K7` y `E3=0` en julio y agosto; localizacion del slot por la clave del
  contrato) y `RevisionPestanasLenteR1T0Plan47Tests` (matrices H1, identidad de canal, espejo local vs
  predicado de produccion, ceros documentados julio/agosto, `Otros`, P3/P5, impacto) → **35/35 verdes**.

### 10.2 Puertas

- **H1→H2: CUMPLIDA** (S2.1..S2.5): fix por regla geometrica con T0 citado; contratos puros en Core;
  frente R1 del 46, `EscritorValoresSlotR1`, `SeccionEspejoR1`, `R1FirmaInterior` y
  `DivergenciasDeclaradasCatalogo` INTACTOS; pestanas fuera de alcance (`SALDOS POR NOTA`, `Recaudo *`,
  BCE, CONSOLIDADO, AJUSTES, VALIDACION_*, GERENTES/REMUNERACION, DetRetri*) con **0 cambios de
  produccion**; `<f>` intacta; build 0 warnings.
- **H2→H3: APROBADA** — corrida fresca desde el OFICIAL a temp (`codigo=0`, sin divergencias);
  **312 celdas** cambiadas SOLO en `Rem. Anticipos R2` (192) y `Reversion Pagos R4` (120), todas
  trazadas a los puntos 1–3 (0 inexplicadas); `<f>` byte-identica (13.000 celdas con formula, 0
  diferencias en 40 hojas); suite **38 Failed / 995 Passed / 1033 Total / 0 Skipped** con set-diff
  **0 altas / 0 bajas** vs `p46-cierre.trx`; Q1 verde; SHA oficial intacto; cero commits; CRLF.

### 10.3 Pendiente (no autorizado aqui)

H3 (docs/doctrina en `.opencode/project-context.md`, estados en `plans/44`, `Docs/Instructivo-Capa-B.md`)
y el **code-review independiente** (lo comisiona el orquestador). La auditoria de la propietaria sigue
PENDIENTE (no se declara paridad/verde por ausencia).

---

## 11. CIERRE DE WARNINGS DEL CODE-REVIEW (2026-10-17) + H3

> El code-review independiente dictamino **APROBADO-CON-OBSERVACIONES** y el Ingeniero ordeno **cerrar
> los warnings** (no dejarlos como deuda). Alcance quirurgico ejecutado y re-verificado.

### 11.1 W-1 — Fin de linea (mecanico)

- `Remuneracion.Core/Models/CanalEfectivoDetalleR2R4.cs` y `CerosDocumentadosDetalleR2R4.cs` pasaron de
  **LF puro** a **CRLF** (contenido intacto, cero bytes de texto cambiados). Verificacion propia:
  `CR == LF` en ambos (139 y 124 lineas, `LF_solo = 0`) y **ningun otro `.cs` del diff en LF**
  (los 7 archivos tocados dan `LF_solo = 0`).

### 11.2 W-2 — Red de seguridad del fail-fast (genero `EscritorValoresSlotR1Tests.S3_EmpresaSinSlot_FailFastSinRemap`)

- `DetalleR2R4EscrituraTests.DetalleR2_FilaSinteticaSinSlot_LanzaErrSinSlot_ConMensajeAccionable`: copia
  los insumos a temp y **inyecta un grupo sintetico** (`D=Vlr Servicio` + cierre `C=Z D=Total`) antes del
  cierre `A=Total` del bloque de ASE2 → la corrida falla con `CodigoError.SinSlot` y mensaje que nombra
  **hoja (`'Rem. Anticipos R2'`) + ASE + nº de fila fuente + firma (`'|||Vlr Servicio'`) + canal (`Z`) +
  bloque**, sin salida parcial (atomicidad).
- `DetalleR2_FilaDocumentada_NoFalla_Y_ConservaSuSlot`: contracara — una fila **dentro** de la lista
  (el `Vlr Intereses` del canal 2 de ASE2, sin slot) no rompe la corrida; el grupo conserva sus valores
  (`E78 = 1.616.173,63`, `E79 = 1.616.580,05`). El test preexistente NO se debilito.

### 11.3 W-3 — Canal compuesto por segundo cierre (la via fuerte)

- **Hueco cerrado:** cuando dos grupos del mismo bloque comparten canal efectivo (R2 ASE5: `E` de ENEL y
  `E` de ENERBIT), el empate se resolvia por **indice de destino** (posicion).
- **Implementacion (Core puro, D-H):** `CanalEfectivoDetalleR2R4.CanalEfectivo` compone
  **`<canal>/<empresa>`** con la empresa del **primer cierre de empresa por debajo**
  (`B=ENEL C=Total`, genero `|ENEL|Total|`) atravesando filas de detalle y otros cierres de canal; el
  mismo predicado corre en fuente y destino. Columnas declaradas por el mapa:
  `ColumnasEtiquetaGrupoR2 = [C]`, `ColumnasSegundoCierreR2 = [B]`, `ColumnasEtiquetaPropiaR2 = [A,B,C]`;
  R4 `ColumnasSegundoCierreR4 = []` (**no compone**: su rotulo de grupo YA es la empresa).
  Ademas, las **filas de cierre de canal** (`C=E D=Total`) toman su canal compuesto (si no, seguian
  empatando por orden).
- **Rotulos derivados de bytes (citados):** `E/ENEL` y `E/ENERBIT` (ASE5, `C388=E`+`B389=ENEL` /
  `C392=E`+`B393=ENERBIT`), `A/EAB` (ASE3 `C166=A`+`B167=EAB`), `9/NUEVO ESQUEMA` (ASE2/ASE4),
  `2/NUEVO ESQUEMA`, `H/OCCIDENTE`, `T/OCCIDENTE`. **Fallback declarado:** si no hay cierre de empresa
  debajo, el canal queda simple (p. ej. la cola del bloque sin empresa posterior).
- **Prueba exigida (el ultimo camino posicional murio), dos niveles:**
  1. `RevisionPestanasLenteR1T0Plan47Tests.T0P47_W3_R2_Ase5_DobleGrupoE_OrdenInvertido_CadaGrupoRecibeLoSuyo`:
     fuente sintetica (mismos bytes reales, grupos `E` invertidos) → cada grupo cae en su slot
     (ENERBIT → 390..393, ENEL → 386..389) y el resto del bloque no se contagia.
  2. `DetalleR2R4EscrituraTests.DetalleR2_Ase5_DobleGrupoE_OrdenInvertido_CadaGrupoRecibeLoSuyo`:
     **end-to-end con la fuente mutada** (contenido de los dos grupos intercambiado en el archivo) → la
     corrida escribe `E386 = 19.713.558,23`, `E388 = 19.732.312,24` (ENEL) y `E390 = 7.522.088,94`,
     `E392 = 7.529.567,64` (ENERBIT) con el orden invertido: **cada grupo recibe lo suyo**.
- **Ambiguiedad:** ningun bloque queda sin desambiguacion; el unico caso sin segundo cierre cae al canal
  simple (declarado arriba). **No hubo STOP.**

### 11.4 S-1 — Evidencia persistida y reconciliacion de la baja heredada

- Set-diff persistido: `Remuneracion.IntegrationTests/TestResults/p47-warnings-setdiff.txt`
  (genero del `p47-h0-setdiff.txt`).
- **Reconciliacion (la narrativa coincide con la evidencia):** linea base del set-diff de FALLOS =
  `p46-cierre.trx` (**38 Failed / 960 Passed / 998 Total / 0 Skipped**); la **unica baja del Plan 46**
  (39 → 38: `ProcesadorPeriodoTests.Ejecutar_Periodo2026072_ConLectorOracle_VeredictosValidacionesNoVacios`,
  ROJO→VERDE) se **hereda** con su causa ya declarada (plans/46 §9.3.3: el test corre Q2 con oraculo y sin
  reconocimiento; pre-fix la corrida bloqueaba con la divergencia declarada OCCIDENTE-ASE5 y no producia
  veredictos; con el espejo R1 consistente cierra y verifica los 5 bloques `VALIDACIONES`, HU-13/2.7).
  El frente 47 **no introduce bajas nuevas** (0 altas / 0 bajas).

### 11.5 Re-verificacion obligatoria (tras W-1..W-3)

| Verificacion | Resultado |
|---|---|
| Build | **0 warnings / 0 errores** |
| Suite | `p47-warnings.trx`: **38 Failed / 999 Passed / 1037 Total / 0 Skipped** (998 + 36 del fixture + 3 de W-2/W-3); set-diff de FALLOS vs `p46-cierre.trx` → **0 altas / 0 bajas** |
| Q1 | `T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback` → **Passed** |
| Corrida fresca | temp desde el OFICIAL + `Docs/Prueba Julio-2/Insumos` (5 ASE) → **codigo=0** |
| Tabla corrida-anterior → nueva | **`p47-h2b` → nueva: 0 celdas `values_changed` en las 40 hojas** (byte-equivalente en valores). El canal compuesto **NO cambia ninguna cifra**: su valor es la **robustez al orden** (W-3). Frente a la corrida 9 se mantiene el cuadro del §12.3 (R2 192 + R4 120) |
| `<f>` | **13.000 celdas con formula, 0 diferencias de texto** (corrida 9 vs nueva, 40 hojas) |
| SHA oficial | `3086AC31…C5214` intacto |
| Proceso | cero commits; CRLF sin mezclar; cero codigo de produccion fuera de R2/R4 |

### 11.6 H3 — Cierre documental

- `.opencode/project-context.md`: nueva doctrina *"detalle R2/R4 por `(firma + canal efectivo compuesto)`
  + ceros documentados + `Otros`"* (solo lo probado).
- `plans/44`: registrado que el **follow-up §1.3 quedo atendido** por el plan hijo 47.
- `Docs/Instructivo-Capa-B.md`: nueva subseccion **R2/R4-detalle** (auditoria fila a fila con las claves
  nuevas + ceros documentados + `ERR-SIN-SLOT`).
- `Docs/T0-Plan47-Revision-Pestanas-LenteR1.md` §12: cerrada la trazabilidad de H1/H2/W-1..S-1.

**FRENTE 47 CERRADO (2026-10-17).** Code-review: APROBADO-CON-OBSERVACIONES → observaciones **cerradas**.
Auditoria de la propietaria: **PENDIENTE** (no se declara paridad/verde por ausencia).

---

## Nota de cierre (Plan 48, 2026-10-19)

El Plan 48 adoptó la **v2** como base única. **Las 9 filas de julio-2 que este plan documentó como
"sin slot" (canal 9 y `Vlr Intereses` del canal 2 de ASE2/ASE4) resultaron FALSO HUECO**: la v2 les da
slot y la brecha de **53.283.379,41** quedó cerrada (ASE2 `E87/E89=831.919,29`; ASE4
`E313/E314/E315=20.152.773,11 / -2.377,86 / 20.150.395,25`). `CerosDocumentados` conserva solo las
**3 claves A/EAB de agosto-2** (re-probadas contra la v2); `Otros → R` y la identidad por canal
compuesto siguen vigentes. No se reescribe la historia; detalle en `plans/48` §9.6 y
`Docs/T0-Plan48-Adopcion-Plantilla-v2.md` §2.2/§2.4.
