---
titulo: Plan 50 - Opcionalidad de fuentes de recaudo y unificacion Q1/Q2
estado: CERRADO 2026-10-21 (A1+A3 implementados y APROBADO CON OBSERVACIONES; A2 declarado como deuda R-MAPA-V2-R2-ASE4; v2 ratificada 0F192AD2…)
fecha: 2026-10-21
numero: 50
planes_ocupados: 01..49 ocupados; este es el 50
padre: reunion con la propietaria 2026-10-21 (doctrina transversal en `.opencode/project-context.md` §"Modelo de la plantilla estandar"; Engram #2124 y ratificacion #2223)
madre_doctrina: plans/20 (V3 y G2-D1 — recorte conciliacion por empresa en Q2) + plans/21 (T0e — invariantes de cierre fail-fast) + plans/47/48 (mapas por empresa, canal efectivo, adopcion v2) + plans/49 (doctrina no-inventar, gate absoluto, deuda declarada; no se toca)
caso_rojo: corrida 2026101 (octubre-1, Q1) aborta en ASE 2 (LIME) — `CalculoInvalidoException — ASE 2: no se encontro la fila Total de la empresa ReciprocidadEaab en la fuente.`
gate_entrada:
  plantilla_v2: NO SE TOCA (SHA `FF5CF7A819110A1E1E2CA7BDAD553DD7A25141ADD472B1904BD496F082642632` intacto antes/despues)
  plantilla_v1: NO SE TOCA (historica; SHA `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` intacto antes/despues)
  alcance_aprobado: A1 (opcionalidad fuentes en mapas por empresa R1/R2/R4) + A2 (unificacion Q1/Q2 — levantar recorte `ProcesadorPeriodo.cs:205-214`) + A3 (control anti-falso-0); prohibido ampliar
  suite_base: p49-h2-suite.trx (1061 Total / 1012 Passed / 49 Failed / 0 Skipped) + lista cerrada del Plan 49
  sin_commits: NUNCA `git commit`/`push`
---

# Plan 50 — Opcionalidad de fuentes de recaudo y unificación Q1/Q2

> SDD completo PROPOSE → DESIGN → SPEC → TASKS. Fuente única de verdad para la implementación una vez aprobado ESTE documento. No se implementa sin aprobación del Ingeniero (punto de bloqueo: sin su OK no se abre H1).
> Lectura obligatoria previa: `.opencode/project-context.md` (doctrina transversal 2026-10-21 + Planes 21/25/26/27/28/29/30/31/33/34/39/40/41/42/45/46/47/48/49) + `plans/20` (V3, G2-D1) + `plans/21` (+ `plans/21 - T0 Evidencia.md`, invariantes T0e) + `plans/47` (detalle R2/R4, ceros documentados) + `plans/48` (v2, §§9.6/9.7) + `plans/49` (gate absoluto, regla celda-referenciada) + `Docs/T0-EXTEMP-Ase5-Plan49.md` (método T0 por bytes, referencia de estilo).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Contexto adoptado (no re-investigar; cada hecho cita su fuente)

### 0.1 El rojo que motiva el plan

Corrida `2026101` (octubre-1, Q1) aborta en ASE 2 (LIME):

```text
CalculoInvalidoException — ASE 2: no se encontró la fila Total de la empresa ReciprocidadEaab en la fuente.
```

Cadena causal (código vigente, verificado en esta planificación):

| Punto | Archivo:línea | Hecho |
|---|---|---|
| Mapa | `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapPorEmpresa.cs:96` | `EditablesR1PorEmpresa[(1,2)]` declara 3 operandos para RECIPROCIDAD en ASE2 (sin marca de opcionalidad) |
| Label | `ExcelDataReaderWorkbookLeafInputReader.cs:1524-1527` (`ObtenerLabelFuente`) | empresa 1 en ASE 2/4 → label `NUEVO ESQUEMA` |
| Búsqueda | `ExcelDataReaderWorkbookLeafInputReader.cs:1407-1416` (`EsFilaEmpresa`, `filasTotal`) | filtra filas `C==label && D==Total` antes de `indiceComponente`; si la fuente no trae la fuente, la lista queda vacía |
| Throw | `ExcelDataReaderWorkbookLeafInputReader.cs:1425-1444` + `:1529-1535` (`ValorDe`) | `filasTotal.ElementAtOrDefault(0)` es null → `CalculoInvalidoException` que nombra ASE + empresa + concepto |
| R4 gemelo | `ExcelDataReaderWorkbookLeafInputReader.cs:1501-1505` (`ExtraerCeldasR4`) | `FirstOrDefault(...) ?? throw` con mensaje `ASE {id}: no se encontró la fila Total de la empresa {nombre} en la fuente R4.` |
| R2 gemelo | `ExcelDataReaderWorkbookLeafInputReader.cs:1473-1483` (`ExtraerCeldasR2`) | `ValorDe(filasEmpresa, 0, ...)` / `ValorDe(filasEmpresa, Count-1, ...)` — misma clase de throw si la lista está vacía |
| Writer | `OpenXmlPlantillaWriter.cs:1281-1321` (`EscribirCeldasConciliacion`) | exige valor por celda (`TryGetValue` + `throw` si falta, mensajes `No hay valor R1/R2/R4 por empresa para {nombre} ({hoja}!{celda}) del ASE {id}`) |
| Recorte Q2 | `Remuneracion.Core/Services/ProcesadorPeriodo.cs:198-215` | `if (!esQuincena2)` omite `leaf.Conciliacion` en Q2 (motivo documentado: HU-20-T0b/G2-D1, divergencia R4-por-empresa Q2 persistente) |

### 0.2 Doctrina de negocio (reunión con la propietaria, 2026-10-21)

Fuente: `.opencode/project-context.md` §"Modelo de la plantilla estándar" + Engram #2124 y ratificación #2223. Se adopta sin re-litigar:

1. **`NUEVO ESQUEMA`, `ENEL`, `OCCIDENTE`, `ENERBIT` son FUENTES DE RECAUDO (columna C)**, no empresas obligatorias. **Pueden no venir** en un período; es normal.
2. **R1, R2 y R4 son reportes independientes**: que una fuente venga en uno no obliga a los otros.
3. **La plantilla estándar manda en la estructura** (fuentes/canales/conceptos por cuadro); **el llenado es dinámico**: se coloca lo que viene; **lo que no viene queda en 0** (la plantilla ya está en 0).
4. **Q1 y Q2 se tratan igual**: no hay reglas particulares por quincena.
5. **Nunca abortar por ausencia de una fuente de recaudo**; registrar la ausencia y continuar.
6. Los **invariantes de cierre de la fuente** (`A='Componente' B='Total'`, `A='Subs/Cont' B='Total'`, `A='Total'`) siguen siendo **fail-fast** si faltan (T0e, Plan 21): son estructura del reporte, no una fuente opcional.

### 0.3 Alcance APROBADO por el Ingeniero (tres puntos, prohibido ampliar)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **A1** | Opcionalidad de las fuentes de recaudo en los mapas por empresa (`EditablesR1PorEmpresa`, `EditablesR2PorEmpresa`, `EditablesR4PorEmpresa`) | Fuente ausente → celda destino en **0** (no aborta), con registro de ausencia auditable. Los **invariantes de cierre** siguen fail-fast. Aplica a R1, R2 y R4. |
| **A2** | Unificación Q1/Q2: la lectura por empresa se ejecuta **igual en Q1 y Q2** | Se levanta el recorte `ProcesadorPeriodo.cs:205-214` (`if (!esQuincena2)`). El DESIGN analiza el resto de divergencias Q2 y decide si A1 las cubre o queda residuo declarado con su T0. **No reescribir mapas a ciegas.** |
| **A3** | Control anti-falso-0: distinguir "fuente ausente" de "fuente presente no reconocida" | Mecanismo mínimo y verificable (p. ej. suma de fuentes presentes vs total del reporte, o vs `Mes/Total` del bloque) + umbral. Si no es viable con la información disponible, se declara con su T0 propio. **No abre workbooks extra ni usa `<v>` de celdas con `<f>` como oráculo.** |

### 0.4 Restricciones duras (vinculantes)

- **Plantillas intocables**: v2 (`Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx`, SHA `FF5CF7…2642`) y v1 histórica (SHA `3086AC31…C5214`). Verificar antes/después.
- **Values-only**: cero cambios de fórmula en plantillas.
- **Tests no se debilitan**: 0 `Skip`, 0 `Assert.True(true)`, 0 borrados. Re-congelamientos solo con evidencia citada y lista cerrada.
- **Build 0 warnings**, CRLF sin mezclar, sin `git commit`/`push`.
- **Mapeo por ETIQUETA** (fuente/canal/concepto) y **columnas por ENCABEZADO**; nunca por posición/ocurrencia.
- **No implementar** el `ConsolidadoAse.Extemp` mal alimentado ni el R2 ASE4 `A/EAB` (deudas ajenas, ya declaradas en Plan 49 §9.7).
- **No tocar** los mapas de `AjustesSfT` que ya tienen `EsOpcional` (`WorkbookLeafCellMapAjustesSfT.cs:120-129,166-177`; consumo en `ExcelDataReaderWorkbookLeafInputReader.cs:395-412,503-521`); son el precedente, no el objeto.

---

## 1. PROPOSE

### 1.1 Intent

Que el motor deje de abortar cuando una **fuente de recaudo** no viene en el período: la ausencia se convierte en **0 explícito con huella auditable** (no en excepción), en **R1, R2 y R4**, en **Q1 y Q2 por igual** — mientras la **estructura rota** (invariante de cierre faltante) sigue abortando en duro nombrando ASE + reporte + fila. Y que esa unificación elimine la causa del recorte Q2 (la conciliación por empresa vuelve a leerse en Q2), con un **control mínimo** que impida que un cambio de etiqueta se convierta en un 0 silencioso (falso verde).

### 1.2 In Scope

- A1: política de ausencia en `ExtraerCeldasR1/R2/R4` (+ `Bloque5Total` R1) con 0 explícito + registro; `CalcularVisibleR1/R2/R4` coherentes con el 0; writer intacto en contrato (recibe el 0 del reader) o con ajuste mínimo decidido en DESIGN §2.2.
- A2: levantar el recorte Q2 en `ProcesadorPeriodo.cs:205-214` + análisis del resto de divergencias Q2 (R4-por-empresa, etc.) con veredicto "A1 lo cubre / residuo declarado con T0".
- A3: señal "fuente presente no reconocida" con datos ya leídos (R1/R2/R4 del mismo ASE) + umbral; si no viable, T0 propio declarado.
- Trazabilidad: mensajes fail-fast residuales (estructura rota) intactos en formato; huella de ausencia en log + resultado.
- Tests: re-congelamiento con lista cerrada vs base `p49-h2-suite.trx`; set-diff; SHAs; CRLF; 0 warnings.

### 1.3 Non-goals / out (vinculante)

- No se corrige el EXTEMP (`ConsolidadoAse.Extemp` mal alimentado, Plan 49 §9.7) ni el R2 ASE4 `A/EAB` (`12.462,26`): frentes ajenos con dueño propio.
- No se tocan plantillas (v1/v2), ni fórmulas, ni sellados de caché, ni `DivergenciasDeclaradasCatalogo`, ni gates del Plan 49.
- No se reescriben mapas por empresa "a ciegas" (G2-D1 sigue vigente como método: toda divergencia Q2 no cubierta por A1 → residuo declarado con su T0, nunca parche).
- No se tocan los mapas `AjustesSfT` con `EsOpcional` (precedente intacto).
- No se implementa nada fuera de A1/A2/A3. Q1 fuera de A1/A2 **como frente de cambio de comportamiento distinto de la ausencia**: Q1 solo cambia en que la ausencia deja de abortar (neutralidad §2.7).
- Commits, UI, CLI, docs de usuario.

### 1.4 Riesgos con mitigación

| ID | Riesgo | Mitigación |
|---|---|---|
| R-FALSO-0 | La fuente **viene** pero el parser no la reconoce (cambio de etiqueta) y A1 escribe 0 silencioso | A3 (control §2.5); tercera forma de etiqueta desconocida ≠ ausencia: si hay filas no reconocidas por encima de umbral → fail-fast, nunca 0 |
| R-INVARIANTE | A1 se come un invariante de cierre y convierte estructura rota en 0 | §2.1: invariantes T0e fuera de la política de ausencia; `ValidarInvariantesEspejoR1` (`ProcesadorPeriodo.cs:623-633`) y `indiceComponente < 0` intactos; tests negativos que lo prueban |
| R-PARCIAL | Ausencia **parcial** (TotalMain presente, TotalMenor ausente) se trata como total y oculta corrupción | T0 decide la granularidad (§4-H0): ausencia total del label vs ausencia de bloque; regla por defecto: bloque intermedio faltante con label presente = fail-fast salvo evidencia T0 que lo declare opcional |
| R-Q2-RESIDUO | A1 no cubre todas las divergencias Q2 y el levantamiento introduce escrituras cruzadas | §2.4: inventario de divergencias Q2; lo no cubierto → residuo declarado con T0 propio, Q2 vuelve a acotarse solo en ese residuo |
| R-Q1 | El cambio altera Q1 más allá de la ausencia (geometría, columnas, escrituras) | G-Q1 (§3-S0): Q1 solo cambia en no-abortar-por-ausencia; set-diff con lista cerrada; ningún movimiento Q1 fuera de lista = STOP |
| R-RUIDO | La huella de ausencia inunda el log y oculta errores reales | §2.3: un `Log.Warning` por (ASE, reporte, fuente) + campo agregado en el resultado; sin volcado por celda |
| R-FIXTURE | Re-congelar tests para "volver a verde" en vez de documentar el nuevo comportamiento | Re-congelamiento solo con evidencia citada (fuente sin la fuente + 0 esperado); lista cerrada; 0 `Skip`/borrados |
| R-CANCEL | A3 compara dos derivaciones del mismo workbook y cancela (lección Plan 49) | Un operando del control es el **total del reporte leído de la fuente** (dominio), no otra celda del workbook (§2.5) |

---

## 2. DESIGN

### 2.1 Dónde vive la opcionalidad y qué forma toma (decisión con trade-offs)

**Hecho de partida:** los tres mapas son `(string Celda, FuenteX Fuente)[]` **sin marca de opcionalidad** (`WorkbookLeafCellMapPorEmpresa.cs:83,133,158`). El throw vive en el **reader** (`ValorDe` `:1529-1535`, `EspecialesDe` `:1537-1543`, `ExtraerCeldasR4` `:1501-1505`), no en el mapa. El writer exige valor por celda (`OpenXmlPlantillaWriter.cs:1287-1316`).

| Candidata | Veredicto | Motivo |
|---|---|---|
| **O-A: flag `EsOpcional` por entrada** (`(Celda, Fuente, bool)[]`, gemelo de `AjustesSfT`) | **Descartada como forma principal** | Los mapas por empresa no declaran "fuentes que pueden faltar por celda" sino operandos de una empresa; la ausencia real es **del label completo en el reporte** (cero filas), no de una celda suelta. Marcar cada tupla obligaría a mantener N flags que siempre valdrían lo mismo por (empresa, reporte, período) y se pudriría con cada forma nueva (D-H). Solo tendría sentido para el caso **parcial** (bloque intermedio), que el T0 debe probar primero si existe. |
| **O-B: set de combinaciones opcionales** (`{(empresa, ASE, reporte)}`) | **Descartada** | Congelaría "quién puede faltar" por período observado (p. ej. `(1,2,R1)` porque octubre-1 no trae `NUEVO ESQUEMA`). Es una lista de ausencias pasadas, no una regla: el próximo período con otra fuente ausente volvería a abortar. Viola la doctrina ("pueden no venir; es normal") y D-H (ramas por empresa/período como dato-duro). |
| **O-C: política general "fuente ausente = 0" con excepción de invariantes** | **Elegida** | La ausencia se detecta por **geometría observada** (cero filas del label en el reporte), no por lista. Cubre cualquier fuente futura sin código nuevo (OCP). Los invariantes de cierre T0e quedan **fuera** de la política por construcción (siguen fail-fast). El T0 solo congela **qué cuenta como "ausente"** (label completo vs bloque parcial) y la evidencia de cada caso. |

**Regla O-C (normativa):**

1. **Nivel de ausencia = label completo por reporte.** Si `filasTotal + filasSubs` (R1), `filasEmpresa` (R2) o la fila `Total` (R4) del label son **cero filas**, la fuente está **ausente**: todos sus operandos del mapa → **0 explícito**, con huella (§2.3). No se llama a `ValorDe`/`EspecialesDe` (no hay throw que capturar: hay rama previa que lo evita).
2. **Bloque parcial con label presente = fail-fast por defecto.** Si el label trae *alguna* fila pero falta el bloque que un operando exige (`TotalMenor`/`TotalMedio`/`SubsMenor`/`Especiales*` sin fila), se mantiene el throw actual (nombra ASE + empresa + concepto). Solo un veredicto T0 con bytes puede mover un caso parcial a "0 documentado" (con motivo + evidencia citada, patrón `CerosDocumentadosDetalleR2R4` del Plan 47 — sin crear lista nueva si no hay caso probado).
3. **`Bloque5Total` (R1, `D=5`):** `FirstOrDefault()` ya devuelve 0 si no hay fila (`:1425-1430`); se mantiene y se le suma la huella (hoy es 0 silencioso: pasa a 0 con huella cuando el label está ausente).
4. **Invariantes intactos:** `indiceComponente < 0` (`:1399-1405`), `ValidarInvariantesEspejoR1` (`ProcesadorPeriodo.cs:623-633`, fuente única `BloqueEspejoAseInputs.FaltantesInvariantesDeCierre`) y el gate de coherencia R1 (`WorkbookLeafCoherence`, si aplica) **no entran** en la política. Mensajes fail-fast residuales intactos en formato (ASE + reporte + fila esperada).
5. **Cero ramas por período/empresa (D-H):** la política no recibe `Periodo`, no compara `empresaId == 1`, no lista labels. El label sigue resolviéndose por `ObtenerLabelFuente` (existente); la ausencia se predica del **resultado observado**, nunca del nombre.

### 2.2 Dónde se materializa el 0 (decisión)

| Candidata | Veredicto | Motivo |
|---|---|---|
| **W-A: el reader devuelve 0 por celda** (`CeldasR1/R2/R4[celda] = 0`) | **Elegida** | El writer (`:1287-1316`) y los visibles (`CalcularVisibleR1/R2/R4`) ya consumen el diccionario por celda; con el 0 explícito **todo el downstream funciona sin ramas**: el writer escribe 0 (values-only, sin `throw`), el visible suma 0, el gate Σ lo ve. El 0 es **auditable** (está en `ConciliacionEmpresaInputs`, no solo en la plantilla). |
| **W-B: el reader omite la celda y el writer deja el 0 de plantilla** | **Descartada** | Exigiría relajar el `throw` del writer (`:1287-1316`) a "si falta, no escribir" — ese `throw` es hoy la red que captura mapas desincronizados; relajarlo lo convierte en silencio para **todos** los casos (incluido bug real). Además el visible calculado desde el diccionario y el 0 de plantilla podrían divergir sin que nadie lo note. |
| Consecuencia | — | `EscribirCeldasConciliacion` **intacto** (sigue exigiendo valor por celda del mapa). El cambio vive en `ExtraerCeldasR1/R2/R4`: ausencia total → poblar las celdas del mapa con `0m` + huella. La plantilla sigue aportando su 0 solo donde el mapa no declara celda (sin cambio). |

### 2.3 Registro de ausencia (qué, dónde, cuánto ruido)

- **Qué:** por cada (ASE, reporte R1/R2/R4, label de fuente) ausente: un `Log.Warning` estructurado con `AseId`, `Reporte`, `Fuente` (label observado, p. ej. `NUEVO ESQUEMA`), `Empresa` y `CeldasEnCero` (conteo, no volcado por celda). Formato administrativo (doctrina Plan 26 D-E: sin "matcher", "finder", "slot", nombres de clase).
- **Dónde (aditivo, sin romper contratos):** campo nuevo en `ConciliacionEmpresaInputs` (p. ej. `FuentesAusentes: IReadOnlyList<string>` con entradas `R1:NUEVO ESQUEMA`, `R4:NUEVO ESQUEMA`) — el modelo es de datos puros, el lector lo puebla, el procesador lo propaga al `ResultadoProcesoPeriodo` (campo aditivo, vacío = comportamiento anterior, patrón Plan 42). El writer no lo consume (values-only intacto). Alternativa "solo log" se descarta: el log no viaja al resultado y la auditoría lo exige durable.
- **Cuánto ruido:** máximo **1 warning por (ASE, reporte, fuente)** por corrida (no por celda, no por bloque). La huella del Plan 42 (reporte durable de declaradas) **no** se reutiliza: la ausencia no es divergencia declarada, es comportamiento normal documentado en el resultado.

### 2.4 Impacto en Q2 (qué se levanta, qué residuo puede quedar)

**Se levanta:** `ProcesadorPeriodo.cs:205-214` — el `if (!esQuincena2)` desaparece; `leaf.Conciliacion = _leafReader.LeerConciliacionEmpresas(...)` corre **igual en Q1 y Q2** (progreso + log actualizados: el mensaje `conciliación por empresa omitida en Q2 (HU-20-T0b...)` se retira; el comentario HU-20-T0b/G2-D1 se sustituye por la referencia a este plan sin reescribir historia).

**Análisis que el T0 + H1 deben cerrar (prohibido reescribir mapas a ciegas):**

| Divergencia Q2 documentada | Estado a la entrada | Veredicto esperado (a confirmar por T0/H1) |
|---|---|---|
| R4-por-empresa Q2 (ASE2-Q2: ENEL+OCCIDENTE, sin `NUEVO ESQUEMA`) — causa del recorte (Plan 20 V3/G2-D1, `ProcesadorPeriodo.cs:199-204`) | Recorte vigente | **Cubierta por A1**: sin la fila, `ExtraerCeldasR4` devuelve 0 con huella en vez de abortar. El mapa HU-08 **no se reescribe**. |
| `leaf.Recaudos` (RESUMEN MES F/G en Q2) | Ya levantado (Plan 20 G2) | Intacto; A1/A2 no lo tocan. |
| Resto de layouts Q2 por empresa (si el T0 encuentra otra divergencia: p. ej. empresa que en Q2 trae bloques parciales distintos) | Desconocido | Si es **ausencia total del label** → cubierta por A1. Si es **parcial con label presente** → fail-fast vigente + **residuo declarado con su T0 propio** (alcance acotado a ese residuo, nunca reescritura del mapa). |
| Escrituras Q2 nuevas | — | Las hojas por empresa (`REMUNERACION_*`, `Recaudo *`) reciben en Q2 los mismos operandos que en Q1 (0 donde ausente). **Ninguna celda fuera de las ya declaradas en los mapas por empresa cambia de valor por A2** (verificación por diff Q2 antes/después salvo los 0 de ausencia). |

### 2.5 A3 viable (el control mínimo)

**Principio (anti-cancelación, Plan 49):** un operando del control es el **total del reporte leído de la fuente** (dominio), no otra celda del workbook.

**Señal propuesta (a verificar por el T0 con bytes):** por (ASE, reporte R1/R2/R4):

```text
Σ(valores de las fuentes PRESENTES según el reader, misma columna que el Total)
vs
Total del reporte (fila invariante de cierre: R1 `A='Componente' B='Total'` col F;
R2 bloque `Total` col E; R4 fila `A='Total'` col D)
```

- Si la diferencia `|Σ_presentes − Total_reporte| > umbral` **y** hay al menos una fuente declarada como ausente en ese (ASE, reporte) → la ausencia es sospechosa: hay dinero no atribuido → **fail-fast** nombrando ASE + reporte + fuentes ausentes + Δ (formato §2.7). Es la firma de "fuente presente no reconocida" (cambio de etiqueta: el dinero está en el total pero en ninguna fuente reconocida).
- Si cierra dentro del umbral → la ausencia es coherente (el total se explica sin la fuente) → se acepta el 0.
- **Umbral:** `±0.5` (tolerancia inviolable del proyecto) **más** la tolerancia de redondeo de la columna si el T0 la prueba necesaria; el T0 congela el umbral final con bytes (julio-2 + agosto-2 + octubre-1). Sin bytes no hay umbral.
- **Dónde vive el juicio:** `ValidadorBasico` (Core, junto a los gates 2.7; patrón Plan 49 §2.1); la **evidencia** (Σ presentes + total del reporte por ASE) la expone el reader en los inputs/snapshot (aditivo, read-only, sin juicio en Infrastructure). Cero workbooks extra (los datos ya están leídos en `LeerConciliacionEmpresas`). **Prohibido `<v>`** de celdas con `<f>` como oráculo (los operandos son literales de fuente, no cachés).
- **Si no es viable:** si el T0 demuestra que el total del reporte no es comparable con Σ presentes en algún reporte (p. ej. R2 con `Especiales` heterogéneos, R4 con signos, columnas no homogéneas), se **declara con T0 propio**: el control se acota a los reportes donde sí cierra, y el resto queda con el **mínimo control posible** (p. ej. "toda fila de la fuente cuyo label no coincide con ningún label conocido + `ObtenerLabelFuente` se reporta como `EtiquetasNoReconocidas` en el resultado; si no está vacía y hay ausencias declaradas → fail-fast"). El T0 decide entre control aritmético vs control de etiquetas, nunca ambos a medias sin veredicto.

### 2.6 Trazabilidad del error residual (qué error queda para "estructura rota")

- **Fuente ausente (caso normal):** no hay error; hay huella (§2.3) + 0.
- **Invariante de cierre faltante:** fail-fast intacto — `ProcesadorPeriodo.ValidarInvariantesEspejoR1` (`:623-633`, mensaje `Espejo estructural: la fuente R1 del ASE {id} ({nombre}) no trae las filas invariantes de cierre en el reporte 'Reporte Componentes R1': {...}`) y `ExtraerCeldasR1:1399-1405` (`no se encontró la fila 'Componente'...`). Para R2/R4: el mensaje de `ValorDe`/`ExtraerCeldasR4` se conserva **solo para el caso parcial** (label presente, bloque faltante) y para label presente sin filas donde el mapa lo exige por estructura (el T0 tipifica cuál es cuál; el formato sigue siendo `ASE {id}: no se encontró la fila {concepto} de la empresa {nombre} en la fuente [R1/R2/R4]`).
- **A3 disparado:** `[ERR-VALIDACION] ASE {n} · {reporte} · control de fuentes: Σ presentes ({s}) vs total del reporte ({t}). Diferencia={d} > ±{umbral} con fuente(s) ausente(s) [{labels}]. (Plan 50 A3: posible fuente presente no reconocida.)` — jerga de código prohibida en el texto (doctrina Plan 26 D-E).

### 2.7 Neutralidad Q1/Q2 (garantía)

- Q1 **solo** cambia en que la ausencia deja de abortar (el rojo 2026101 pasa a 0 con huella). Ninguna geometría, columna, fórmula, sello o gate cambia por A1/A2 fuera de ese caso.
- Q2 **no gana escrituras nuevas** fuera de la conciliación por empresa (que es exactamente lo que el recorte omitía). Verificación: diff de workbook Q2 antes/después con solo deltas trazados a ausencias (build + goldens + e2e Q1/Q2 en §3-S0).
- `AjustesSfT = 0` en Q1, `RESUMEN MES` D/E-vs-F/G por dominio, DetRetri-vs-R10 y sello de fechas: intactos.

### 2.8 Arquitectura del cambio (contract-first)

- **Core (puro):** política O-C como predicado puro donde el T0 la sitúe (p. ej. helper `FuenteAusenteR1/R2/R4` o rama previa en el reader existente — H1 decide citando; D-H intacto: cero literales de label/período en producción); campo aditivo `FuentesAusentes` en `ConciliacionEmpresaInputs` (+ propagación aditiva en `ResultadoProcesoPeriodo`); juicio A3 en `ValidadorBasico` (método nuevo junto al gate 2.7, mensaje §2.6).
- **Infrastructure (lectura + values-only):** `ExtraerCeldasR1/R2/R4` implementan la rama de ausencia total (0 explícito + huella); `Bloque5Total` suma huella; `EscribirCeldasConciliacion` intacto (sigue exigiendo valor por celda); `ProcesadorPeriodo` levanta el recorte (§2.4).
- **Tests:** T0 read-only (bytes) + tests de ausencia (sintéticos donde se pueda, con fuentes reales donde el proyecto lo exige) + re-congelamiento con lista cerrada + set-diff.
- **Docs:** doctrina en `.opencode/project-context.md` (solo lo probado) + nota en `plans/20` (recorte levantado por este plan) sin reescribir historia.

### 2.9 Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Reader = evidencia (detecta ausencia, devuelve 0, reporta huella) / Validador = juicio (A3 + gates 2.7) / Mapas = datos (sin flags de ausencia). El writer no juzga (sigue exigiendo valor por celda). |
| **OCP** | ✅ Fuente nueva o ausente nueva = cero código (la política la absorbe por geometría observada). Solo un caso parcial probado por T0 añade comportamiento, como dato citado, nunca como rama. |
| **DIP** | ✅ Cambios vía interfaces existentes (`IWorkbookLeafInputReader`, `IValidador`, `IPlantillaWriter`); `ProcesadorPeriodo` sigue dependiendo de abstracciones; el campo de huella es aditivo. |
| **Best practices** | ✅ Contract-first (T0 congela qué cuenta como ausente + umbral A3 antes del juicio); fail-fast accionable (ASE + reporte + fuente/fila + Δ); gates que leen lo escrito + dominio (nunca `<v>` stale); tolerancia ±0.5; values-only; D-H; CRLF; sin commits. |
| **Performance** | ✅ O(filas del reporte) ya leídas; sin workbooks extra; el control A3 es O(ASE × reporte) sobre valores en memoria. Tests de ausencia sintéticos donde sea posible (sin disco) + e2e con fuentes reales donde el proyecto lo exige. |

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda T)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa con trx + **set-diff vs línea base `p49-h2-suite.trx` (1061/1012/49/0)**: altas/bajas una por una con causa citada. **Rojo nuevo fuera de lista = STOP con causa.** 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados.
- G-SHA: v1 `3086AC31…C5214` + v2 `FF5CF7…2632` intactos antes/después.
- G-PROC: cero commits; CRLF sin mezclar en lo tocado.
- G-D-H: cero literales/ramas por período o empresa en producción (grep de cierre; los labels son datos del catálogo existente, no ramas nuevas).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale) o la fuente ya leída; **cero workbooks extra abiertos por A3**.
- G-FORMULA: `<f>` byte-idéntica en todo el diff (cero `new CellFormula`, cero `InsertRow`/`DeleteRow`).
- G-TRAZA: cada error nuevo cumple el formato §2.6 (ASE + reporte + fuente/fila + esperado vs obtenido + Δ donde aplique).
- G-Q1: Q1 solo cambia en no-abortar-por-ausencia; **cualquier movimiento Q1 fuera de la lista cerrada = STOP con causa byte-probada**.

### S1 — Criterios H0/T0 (read-only, cero código de producción)

- S1.1: inventario de ausencia por (ASE × reporte × label) en octubre-1 (Q1 2026101) + julio-2 + agosto-2: qué labels faltan por reporte, con conteo de filas observadas (0 = ausente total) vs bloques parciales.
- S1.2: veredicto de granularidad (total vs parcial, §2.1 puntos 1-2) con bytes: qué casos son ausencia total (→ 0) y cuáles son parciales (→ fail-fast salvo prueba que los mueva a 0 documentado).
- S1.3: veredicto A3 con bytes: Σ presentes vs total del reporte por (ASE, reporte) en las tres muestras; umbral congelado (base ±0.5 + ajuste probado si aplica) o declaración de no-viabilidad con T0 propio y control mínimo alternativo (§2.5).
- S1.4: inventario de divergencias Q2 restantes (R4-por-empresa + resto de §2.4) con veredicto "A1 lo cubre / residuo con T0 propio" por ítem.
- S1.5: matriz de mensajes residuales (qué throw queda para cada caso de estructura rota) con formato §2.6.
- S1.6: fixture read-only verde (BCL ZIP+XML, cero producción) que congela S1.1-S1.5.

### S2 — Criterios H1 (A1 + A2, solo lo probado por el T0)

- S2.1: ausencia total del label → 0 explícito por celda del mapa + huella (§2.3); invariantes T0e y parciales intactos en fail-fast.
- S2.2: recorte Q2 levantado (`ProcesadorPeriodo.cs:205-214` sin `if (!esQuincena2)` para conciliación); `leaf.Conciliacion` poblada en Q1 y Q2; mensaje HU-20-T0b retirado/sustituido con cita a este plan.
- S2.3: reproducción del rojo 2026101: ASE2 `NUEVO ESQUEMA` ausente → corrida continúa con 0 + huella (no `CalculoInvalidoException`).
- S2.4: build 0 warnings + greps D-H + G-FORMULA + G-SHA + G-Q1.
- S2.5: `CalcularVisibleR1/R2/R4` coherentes con el 0 (el visible de una fuente ausente es 0; el gate Σ lo ve).

### S3 — Criterios H2 (A3 + regresión)

- S3.1: control A3 implementado (juicio en `ValidadorBasico`, evidencia aditiva del reader) con umbral T0; dispara con el Δ exacto ante etiqueta cambiada simulada y no dispara en corridas sanas con ausencia coherente.
- S3.2: e2e Q1 (2026101) + Q2 (julio-2/agosto-2) con ausencias: tabla corrida→nueva justificada por celda; cambios solo trazados a ausencias (0 con huella).
- S3.3: set-diff cumple S0 (altas == lista cerrada citada; cero fuera de lista) + SHAs intactos + code-review independiente (spec compliance + calidad).
- S3.4: si el T0 declaró A3 no-viable en algún reporte: control mínimo alternativo implementado + T0 propio registrado como follow-up con dueño.

### S4 — Criterios H3 (docs/cierre)

- S4.1: doctrina del plan en `.opencode/project-context.md` (solo lo probado: opcionalidad de fuentes, unificación Q1/Q2, control A3 + umbral).
- S4.2: nota en `plans/20` (recorte G2-D1 levantado por este plan) sin reescribir historia; desviaciones con fecha/motivo/evidencia.
- S4.3: diff de docs citado; cero código de producción en H3.

---

## 4. TASKS (fases H0..H3 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero código de producción)

| ID | Tarea | Archivo:línea (referencia) | Verificación |
|---|---|---|---|
| H0.1 | Inventario de ausencia por bytes (BCL ZIP+XML): por (ASE × R1/R2/R4 × label) en 2026101 + julio-2 + agosto-2 — filas observadas del label (0 = ausente total) vs bloques parciales; labels `NUEVO ESQUEMA`/`ENEL`/`OCCIDENTE`/`ENERBIT` por reporte | Fuentes en `Docs/Insumos/` + período 2026101 del Ingeniero; lectores `ExcelDataReaderWorkbookLeafInputReader.cs:1386-1535` como guía de qué filas buscar | Tabla S1.1 publicada con conteos |
| H0.2 | Veredicto de granularidad total-vs-parcial (§2.1): para cada caso H0.1, clasificar ausencia total (→ 0) o parcial con label presente (→ fail-fast salvo prueba); congelar qué cuenta como "ausente" | `ExcelDataReaderWorkbookLeafInputReader.cs:1411-1444,1467-1489,1500-1519` | Veredicto S1.2 escrito; parciales listados con su throw residual |
| H0.3 | Veredicto A3 por bytes: Σ presentes vs total del reporte por (ASE, reporte) en las tres muestras; congelar umbral (base ±0.5) o declarar no-viabilidad con control mínimo alternativo | Invariantes T0e (`BloqueEspejoAseInputs`, `ProcesadorPeriodo.cs:623-633`); totales `Componente/Total`, `Subs/Cont/Total`, `Total` | Veredicto S1.3 + umbral citado, o T0 propio declarado |
| H0.4 | Inventario de divergencias Q2 restantes (§2.4): R4-por-empresa + resto de layouts por empresa en Q2; veredicto por ítem "A1 lo cubre / residuo con T0 propio" | `ProcesadorPeriodo.cs:198-215`; `WorkbookLeafCellMapPorEmpresa.cs:83-177` | Tabla S1.4; sin veredicto no se abre H1 |
| H0.5 | Congelar fixture read-only nuevo (BCL ZIP+XML, cero producción; gemelo de `ContratoSeccionPeriodFreeT0Plan46Tests` / `RevisionPestanasLenteR1T0Plan47Tests`) con S1.1-S1.5 | `Remuneracion.IntegrationTests/` (nuevo `FuentesRecaudoOpcionalesT0Plan50Tests.cs`) | Fixture verde; S1.6 |
- **Puerta H0→H1:** fixture verde + veredictos S1.1-S1.5. Si H0 demuestra forma no absorbible sin ramas → STOP declarado. Sin T0 verde no hay H1.

### H1/F1 — A1 + A2 (solo donde el T0 probó ausencia)

| ID | Tarea | Archivo:línea | Verificación |
|---|---|---|---|
| H1.1 | Rama de ausencia total en `ExtraerCeldasR1`: si `filasTotal + filasSubs` del label están vacías → poblar las celdas del mapa con `0m` + huella; si hay filas pero falta un bloque exigido → throw vigente | `Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs:1386-1452` | S2.1 (R1); repro 2026101 ASE2 sin throw |
| H1.2 | Ídem en `ExtraerCeldasR2` (vacía `filasEmpresa` → 0 + huella; `Especiales` ausente por encabezado ya es 0 — sumar huella) | `ExcelDataReaderWorkbookLeafInputReader.cs:1454-1489` | S2.1 (R2) |
| H1.3 | Ídem en `ExtraerCeldasR4` (sin fila `Total` del label → 0 + huella en vez de `?? throw` `:1501-1505`); `Bloque5Total` suma huella cuando el label está ausente | `ExcelDataReaderWorkbookLeafInputReader.cs:1491-1535` | S2.1 (R4) |
| H1.4 | Campo aditivo `FuentesAusentes` en `ConciliacionEmpresaInputs` (+ propagación aditiva en `ResultadoProcesoPeriodo`); `CalcularVisibleR1/R2/R4` verificados con el 0 (sin ramas nuevas) | `Remuneracion.Core/Models/ConciliacionEmpresaInputs.cs:12-53`; `ExcelDataReaderWorkbookLeafInputReader.cs:1548-1581`; `ResultadoRemuneracion.cs` / `ResultadoProcesoPeriodo` | S2.5; vacío = comportamiento anterior |
| H1.5 | Levantar recorte Q2: eliminar `if (!esQuincena2)` de conciliación; actualizar progreso/log/comentario HU-20-T0b con cita a este plan | `Remuneracion.Core/Services/ProcesadorPeriodo.cs:198-215` | S2.2; `leaf.Conciliacion` poblada en Q1 y Q2 |
| H1.6 | Tests focalizados de ausencia (rojo→verde honesto): fuente ausente → 0 + huella; invariante faltante → fail-fast intacto; parcial con label presente → fail-fast vigente | `Remuneracion.IntegrationTests/` (nuevos `FuentesAusentesPlan50Tests.cs`) | S2.1-S2.3; build 0 warnings; greps D-H |
- **Puerta H1→H2:** S2.1-S2.5 verdes. Sin esto no se abre H2.

### H2/F2 — A3 + regresión

| ID | Tarea | Archivo:línea | Verificación |
|---|---|---|---|
| H2.1 | Evidencia A3 (aditiva, read-only): Σ presentes + total del reporte por (ASE, reporte) desde lo ya leído; sin workbooks extra; sin `<v>` | `ExcelDataReaderWorkbookLeafInputReader.cs` (evidencia) + `ValidacionCruzadaSnapshot.cs` / inputs (campo aditivo) | S3.1 (evidencia) |
| H2.2 | Juicio A3 en `ValidadorBasico` (método nuevo junto al gate 2.7, mensaje §2.6); `ClasificarDivergenciasDeclaradas` lo hereda (última puerta, Plan 42 D-F: A3 disparado = NO declarada, falla en duro) | `Remuneracion.Core/Services/ValidadorBasico.cs:472-578` (zona) | S3.1; caso etiqueta-cambiada dispara con Δ exacto; caso sano no dispara |
| H2.3 | Test de simulación A3 (sintético donde sea posible + e2e con fuente mutada donde el proyecto lo exige): etiqueta cambiada → fail-fast A3; ausencia coherente → verde con huella | `Remuneracion.IntegrationTests/` (nuevo `ControlAntiFalsoCeroPlan50Tests.cs`) | S3.1; 0 `Skip` |
| H2.4 | Regresión e2e Q1 (2026101) + Q2 (julio-2/agosto-2): tabla corrida→nueva justificada por celda; suite trx + set-diff vs base `p49-h2-suite.trx` (1061/1012/49/0) con lista cerrada; SHAs; G-Q1 | `Remuneracion.IntegrationTests/TestResults/p50-h2-suite.trx` + `p50-h2-setdiff.txt` | S3.2-S3.3 |
| H2.5 | Code-review independiente (spec compliance + calidad): trazabilidad cada cambio → A1/A2/A3; ningún test debilitado; CRLF; SHAs | — | Puerta H2→H3: review APROBADO (o APROBADO-CON-OBSERVACIONES cerradas) |
- **Puerta H2→H3:** set-diff S0 + G-Q1 + SHAs + review aprobado.

### H3/F3 — Docs/cierre

| ID | Tarea | Archivo:línea | Verificación |
|---|---|---|---|
| H3.1 | Doctrina en `.opencode/project-context.md` (opcionalidad de fuentes + unificación Q1/Q2 + control A3 y umbral; solo lo probado) | `.opencode/project-context.md` | Diff citado; S4.1 |
| H3.2 | Nota en `plans/20` (recorte G2-D1 levantado por el Plan 50) sin reescribir historia; desviaciones del plan con fecha/motivo/evidencia | `plans/20 - conciliaciones-resumen-mes-q2-r10-remuneracion.md` | Diff citado; S4.2 |
| H3.3 | `Docs/Instructivo-Capa-B.md` si aplica (qué significa la huella de ausencia para el administrativo; que un 0 con huella es normal y un A3 es defecto a escalar) | `Docs/Instructivo-Capa-B.md` | S4.3 |
- **Puerta H3:** diff de docs citado; cero código de producción en H3.

**Orden sugerido:** H0 → H1 → H2 → H3. Regresión continua (suite + SHAs + G-Q1 en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

---

## 5. Riesgos (tabla consolidada + gates por fase)

Ver §1.4 (R-FALSO-0, R-INVARIANTE, R-PARCIAL, R-Q2-RESIDUO, R-Q1, R-RUIDO, R-FIXTURE, R-CANCEL). Gates por fase:

- **Gate H0→H1:** fixture T0 verde + veredictos S1.1-S1.5. Sin T0 verde no hay H1; si H0 declara STOP, no se abre H1.
- **Gate H1→H2:** S2 verde (solo ausencias probadas) + build 0w + greps D-H + G-FORMULA + G-SHA + G-Q1.
- **Gate H2→H3:** set-diff S0 (altas == lista cerrada) + Q1 intacto + SHAs intactos + tabla e2e justificada + review aprobado.
- **Gate H3:** diff de docs citado; cero código de producción en H3.

---

## 6. Architecture Validation Certificate

- **SOLID:** SRP (reader = evidencia + 0 explícito / validador = juicio A3 + gates / mapas = datos sin flags; writer no juzga); OCP (fuente ausente nueva = cero código; parcial nuevo = fila citada en T0, cero ramas); LSP/DIP (dependencias por interfaces existentes; campos aditivos, contratos intactos); ISP (contratos `IRecaudoReader`/`IWorkbookLeafInputReader`/`IValidador`/`IPlantillaWriter` intactos).
- **Best Practices:** contract-first (H0 congela ausencia + umbral antes del juicio); fail-fast accionable (ASE + reporte + fuente/fila + Δ); gates que leen workbook + dominio (nunca `<v>` stale); tolerancia ±0.5; values-only; D-H; CRLF; sin commits.
- **Performance:** O(filas ya leídas) + O(ASE × reporte) en memoria; sin workbooks extra; sin re-lectura por celda; regens a temp fresca.

---

## 7. Clarification Gate (requiere aprobación del Ingeniero; sin esto no se abre H1)

1. **O-C + rama previa (no flag por celda, no set de combinaciones):** ratificar la política general "fuente ausente = 0" con excepción de invariantes (§2.1) frente a las alternativas O-A/O-B.
2. **W-A (reader devuelve 0; writer intacto):** ratificar que el 0 se materializa en el reader y el `throw` del writer (`OpenXmlPlantillaWriter.cs:1287-1316`) sigue exigiendo valor por celda, frente a la alternativa W-B (omitir y dejar el 0 de plantilla).
3. **Granularidad:** ratificar que el bloque parcial con label presente sigue en fail-fast por defecto y solo un T0 con bytes lo mueve a 0 documentado (§2.1 punto 2).
4. **A3 y umbral:** ratificar la señal Σ-presentes-vs-total-del-reporte + umbral base ±0.5 (§2.5) frente a "solo control de etiquetas"; y que si el T0 la declara no-viable se implementa el mínimo alternativo con su T0 propio.
5. **Q2 y lista cerrada:** ratificar el levantamiento del recorte (§2.4) y que todo residuo Q2 no cubierto por A1 queda declarado con su T0 (nunca parche al mapa); ratificar la base `p49-h2-suite.trx` (1061/1012/49/0) para el set-diff y que rojo fuera de lista = STOP.
6. **Puerta:** sin 1–5 no se abre H1; sin H1 verde no hay H2; sin review no hay H3.

Sin estas aprobaciones el plan no avanza de H0 a H1 (1, 2, 4, 5) ni cierra H2 (toda la regresión + §S3) ni H3 (4, 5).

**RATIFICACION 2026-10-21 (Ingeniero) — APROBADO 5/5:** (1) O-C: política general "fuente ausente = 0" con excepción de invariantes; (2) W-A: el reader devuelve 0 y el writer queda intacto; (3) parcial con label presente = fail-fast por defecto (solo un T0 lo mueve a 0); (4) A3 con Σ-presentes-vs-total ±0.5 y mínimo alternativo si el T0 la declara no viable; (5) levantamiento del recorte Q2, residuos declarados con T0 y set-diff contra `p49-h2-suite.trx`. **H0 HABILITADO.**

---

## 8. Estado

**APROBADO 2026-10-21.** Clarification Gate §7 ratificado 5/5 por el Ingeniero. **H0 habilitado:** T0 read-only bloqueante (congela ausencias reales, umbral A3 y residuo Q2); sin H0 verde no se abre H1.

### Estado de implementación (2026-10-21, H0→H3)

| Fase | Estado | Evidencia |
|---|---|---|
| **H0** (T0 read-only) | **VERDE con residuos declarados** | `Docs/T0-Plan50-FuentesRecaudoOpcionales.md` · fixture `FuentesRecaudoOpcionalesT0Plan50Tests` (68 verdes) |
| **H1** A1 (opcionalidad) | **DONE** | `ExcelDataReaderWorkbookLeafInputReader.ExtraerCeldasR1/R2/R4` + `FuentesAusentes` (reader → `ConciliacionEmpresaInputs` → `ResultadoProcesoPeriodo`) · `FuentesAusentesPlan50Tests` (7 verdes) |
| **H1** A2 (unificación Q1/Q2) | **BLOQUEADO — recorte MANTENIDO** | `R-MAPA-V2-R2-ASE4`: `Rem. Anticipos R2!E285` es **fórmula** en la v2 (`E192+E203-K192`; v1: literal 0) → `ERR-PLANTILLA` al escribir. Reactivar exige re-congelar el mapa R2 (1,4) por T0 (clase Plan 48). T0 §4.3 |
| **H2** A3 (anti-falso-0) | **DONE (control mínimo alternativo)** | Aritmética Σ-presentes-vs-total **no viable** en R1/R2 (T0 §3) → control de etiquetas: `ConciliacionEmpresaInputs.EtiquetasNoReconocidas` (reader) + `ValidadorBasico.ValidarFuentesRecaudoNoReconocidas` (juicio) · `ControlAntiFalsoCeroPlan50Tests` (5 verdes) |
| **H2** regresión | **VERDE sin altas** | `TestResults/p50-h2-suite.trx` = **1141/1074/67/0** vs base viva `p50-h0-baseline.trx` = 1061/994/67/0 → **0 altas, 0 bajas, +80 tests verdes** (`TestResults/p50-h2-setdiff.txt`) |
| **H3** docs | **DONE** | `.opencode/project-context.md` (doctrina Plan 50) · `plans/20` (nota de vigencia del recorte G2-D1) · `Docs/Instructivo-Capa-B.md` §11 (0 con huella = normal; etiqueta no reconocida = escalar) |

**Residuos declarados con su propio T0 (no se parchean):**

1. `R-PARCIAL-OCT-ASE3-R1-OCC` — **RECLASIFICADO 2026-10-22:** NO es estructura rota. La fuente OCCIDENTE no trae la primera sección (`AFaseo`) en octubre-1, y eso es normal (fuentes Y canales pueden faltar; indicación de la propietaria). El defecto es el **mapa por ocurrencia**: debe resolver por sección y dejar en 0 la sección ausente, sin reutilizar `filas[count-1]` (eso duplicaría el bloque grande). Fix en plan propio.
2. `R-SIGMA-Q2` — el gate Σ empresas-vs-bloque no descompone el visible de bloque en Q2 (14 divergencias medidas) → acotado a Q1 (comportamiento Q2 idéntico al anterior). T0 §4.2.
3. `R-MAPA-V2-R2-ASE4` — celda destino R2 (1,4) es fórmula en la v2 → A2 no se habilita. T0 §4.3.

**Nota de línea base (G-SHA):** el commit `d55e621` reemplazó el binario de la v2 (`FF5CF7A8…` → `0F192AD2…`) **antes** de este plan; el set-diff se hizo contra una base re-medida en el árbol actual para no atribuir al Plan 50 esos 18 rojos pre-existentes. La v2 **no** fue modificada por este plan (SHA igual antes/después: `0F192AD2…`) y la v1 sigue intacta (`3086AC31…C5214`).

### Cierre (2026-10-21)

- **Code-review independiente:** **APROBADO CON OBSERVACIONES** (0 bloqueantes). Build 0 warnings re-ejecutado; suite `1141/1074/67/0` reproducida; set-diff vs base viva: 0 altas / 0 bajas; las 2 bajas (`ValidadorTotalesR1GateTests.S1_Agosto2026082_GateVerde`, `S1_Julio2026072_GateVerde`) son **verdes legítimos** (verificados en vivo, gates no tocados por el Plan 50); los 20 rojos son preexistentes por `d55e621` y no bloquean este plan.
- **W-1/W-2 corregidos:** `plans/20` y `Docs/Instructivo-Capa-B.md` normalizados a CRLF sin BOM; `git diff --check` vacío.
- **Decisión del Ingeniero (2026-10-21):** ratificar la v2 `0F192AD2…` como nueva base; cerrar el Plan 50 con A1+A3; **A2 queda como deuda declarada** con T0 propio del mapa R2 (1,4).
- **A2 (deuda):** `R-MAPA-V2-R2-ASE4` — `Rem. Anticipos R2!E285/K285` son fórmula en la v2; habilitar Q2 exige re-congelar el mapa R2 (1,4) por T0 (clase Plan 48).
- **Pendientes de otros frentes:** re-congelación de los 20 rojos de la v2 `d55e621` (T0 propio); `ReportePagosxBanco` mapa posicional (HU-09); `R-PARCIAL-OCT-ASE3-R1-OCC`; `R-SIGMA-Q2`.
- **Estado: CERRADO 2026-10-21** (A1+A3 verdes; A2 y residuos declarados con dueño).
