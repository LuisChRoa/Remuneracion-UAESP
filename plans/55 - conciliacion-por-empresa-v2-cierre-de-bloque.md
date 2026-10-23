---
titulo: Plan 55 - Conciliacion por empresa v2, cierre de bloque (las 5 deltas del gate SigmaEmpresas)
estado: CERRADO 2026-10-23 (H0→H1→H2→H3; code-review APROBADO 0 bloqueantes; p55-H3-suite.trx 1484/1417/67/0 con set-diff de rojos vacio; CLI 2026101 sin ERR-VALIDACION y DetRetri = R10 en los 5 ASE)
enmienda_dinamica: 2026-10-23 — doctrina de la propietaria de la plantilla estandar, ratificada por el Ingeniero ("todo debe quedar dinamico, nada quemado por ASE o situaciones particulares"). Invalida el mecanismo de tabla congelada por (empresa, ASE) (D4/H1 originales: ni parchada con tres entradas). Reescribe DESIGN y H1; conserva H0/T0 integro como evidencia (Docs/T0-Plan55-Conciliacion-Por-Empresa.md, fixture ConciliacionEmpresaCierreBloqueT0Tests 81/81, p55-H0-suite.trx 1455/1388/67/0, mismos 67 rojos).
fecha: 2026-10-23
numero: 55
planes_ocupados: 01..54 ocupados; este es el 55
padre: plans/54 - contrato-unico-lectura-r1-q1-q2.md
madre_doctrina: .opencode/project-context.md (Plan 50 S0.2 doctrina fuentes opcionales + invariantes fail-fast) + plans/50 (S0.2 y S8, R-MAPA-V2-R2-ASE4) + plans/53 (mapeo por seccion por operando) + plans/54 (lectura R1 unica)
caso_rojo: corrida CLI real 2026101 Q1 plantilla v2 --cinco-ase — RESULTADO ERROR codigo=ERR-VALIDACION salida=1, 5 items ValidarSigmaEmpresas (ASE1-R1, ASE1-R4, ASE2-R2, ASE2-R4, ASE4-R1)
gate_entrada:
  plantilla_v2: NO SE TOCA (SHA 0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB — verificar intacto antes/despues)
  plantilla_v1: NO SE TOCA (historica; SHA 3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214 intacto antes/despues)
  alcance_aprobado: T0 por bytes primero (H0 read-only); cero codigo de produccion en H0; prohibido ampliar sin nuevo OK del Ingeniero
  suite_base: p53-h2-suite.trx (1289/1222/67/0), estado post-Plan 54 (p54-H3-suite.trx, 1374/1307/67/0, mismo conjunto de rojos) y estado post-H0 de este plan (p55-H0-suite.trx, 1455/1388/67/0: +81 altas del fixture T0, 0 bajas, mismos 67 rojos)
  sin_commits: NUNCA git commit / git push
---

# Plan 55 — Conciliacion por empresa v2: cierre de bloque (SigmaEmpresas 2026101)

> SDD completo PROPOSE → DESIGN → SPEC → TASKS. Fuente unica de verdad para la implementacion una vez aprobado ESTE documento. No se implementa sin aprobacion del Ingeniero (punto de bloqueo: sin su OK no se abre H0; estado inicial PROPUESTO).
> Lectura obligatoria previa: `plans/54` (S7 residuos incl. R-CIERRE-EMPRESAS-2026101, S8) + `plans/50` (S0.2 doctrina, S8 residuo R-MAPA-V2-R2-ASE4) + `Docs/T0-Plan50-FuentesRecaudoOpcionales.md` S4.3 + `plans/53` (contrato por operando) + `plans/52` + `plans/46` + `plans/48` + `.opencode/project-context.md`.
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Espanol; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Contexto adoptado (confirmado; no re-investigar de cero, solo confirmar por bytes en el T0)

### 0.1 De donde viene esto

El Plan 54 (`plans/54 - contrato-unico-lectura-r1-q1-q2.md`, CERRADO 2026-10-22) unifico la lectura R1 y elimino la doble resolucion por quincena. Con eso la corrida real `2026101` **ya no aborta en R1** y los 5 `DetRetri` esperados coinciden al peso con el R10 oficial (`14.842.425.136` / `11.302.637.605` / `8.823.787.300` / `10.483.172.762` / `7.550.668.665`; Plan 54 S8). Recien entonces se alcanza una validacion que antes era inalcanzable y que **es el objeto de este plan**.

### 0.2 El rojo que motiva el plan (verbatim de la corrida real)

Corrida CLI real (`--periodo 2026101 --carpeta "Docs/Prueba Octubre-1/Insumos" --plantilla "<v2>" --cinco-ase`):

```text
RESULTADO ERROR codigo=ERR-VALIDACION salida=1
[ERR-VALIDACION] ASE 1: S visibles de empresas en R1 (14815023989.82) no coincide con el visible de bloque (14815021974.75). Detalle: ReciprocidadEaab=0, Enel=14471401705.97, Enerbit=0, OccidenteDirecta=343622283.85, EaabCiudadLimpia=0
[ERR-VALIDACION] ASE 1: S visibles de empresas en R4 (-134910) vs visible de bloque (-3833780)
[ERR-VALIDACION] ASE 2: S visibles de empresas en R2 (95504450.32) vs visible de bloque (95634974.32)
[ERR-VALIDACION] ASE 2: S visibles de empresas en R4 (-20858585.45) vs visible de bloque (-21803648.45)
[ERR-VALIDACION] ASE 4: S visibles de empresas en R1 (10474277163.42) vs visible de bloque (10473967813.42). Detalle: ReciprocidadEaab=309350.00, Enel=9882391541.17, Enerbit=0, OccidenteDirecta=591576272.25, EaabCiudadLimpia=0
```

Deltas medidas (valor absoluto, para clasificar en el T0):

| # | ASE × reporte | S empresas | Visible bloque | Delta |
|---|---|---:|---:|---:|
| D1 | ASE1 R1 | 14.815.023.989,82 | 14.815.021.974,75 | **2.015,07** (= \|primera fila `Mes/Total` de la fuente, que vale −2.015,07\|) |
| D2 | ASE4 R1 | 10.474.277.163,42 | 10.473.967.813,42 | **309.350,00** (= **exactamente** `ReciprocidadEaab`) |
| D3 | ASE2 R2 | 95.504.450,32 | 95.634.974,32 | **130.524,00** |
| D4 | ASE1 R4 | −134.910,00 | −3.833.780,00 | **3.698.870,00** |
| D5 | ASE2 R4 | −20.858.585,45 | −21.803.648,45 | **945.063,00** |

### 0.3 Hechos verificados (citarlos; el T0 los confirma por bytes, no los re-investiga)

| # | Hecho | Cita |
|---|---|---|
| F1 | El gate es `WorkbookLeafCoherence.ValidarSigmaEmpresas` (`:93-121`): por ASE compara `S Conciliacion.VisibleR1` vs `leaf.R1.TotalOportunoEsperadoPorAse` (`:105-112`), `S VisibleR2` vs `leaf.R2.TotalOportunoEsperado` (`:113-116`), `S VisibleR4` vs `leaf.R4.TotalReversionEsperada` (`:117-120`). Tolerancia ±0,5 (`AsegurarDentroDeTolerancia :123-131`). **Solo corre en Q1** (`if (conciliacion.Count == 0 \|\| leaf.Periodo.NumeroQuincena != 1) return;`, linea 100) | `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs:93-131` |
| F2 | El gate se invoca desde `OpenXmlPlantillaWriter.GenerarWorkbook` (`:236-239` valida archivo/rutas, `:241` coherencia multi-ASE, `:243-247` loop `ValidarSigmaEmpresas(leaf.Conciliacion, leaf)` por leaf antes de escribir) | `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs:236-247` |
| F3 | El agregado por empresa vive en `ExcelDataReaderWorkbookLeafInputReader.ExtraerCeldasR1/ExtraerCeldasR2/ExtraerCeldasR4` + `LeerConciliacionEmpresas`, con el mapa `WorkbookLeafCellMapPorEmpresa` (`EditablesR1PorEmpresa :92`, `EditablesR2PorEmpresa :213`, `EditablesR4PorEmpresa :238` — R1 ya con tupla `(Celda, FuenteR1, SeccionEsperadaR1)` del Plan 53), el label por empresa (`ObtenerLabelFuente`) y la sinonimia (`SinonimosEmpresaR1.SonMismaEmpresa`, Plan 34) | archivos citados + Plan 53/Plan 34 |
| F4 | **Tabla de direcciones desalineada de la v2 — prueba documentada:** `R-MAPA-V2-R2-ASE4` (Plan 50 S8 y `Docs/T0-Plan50-FuentesRecaudoOpcionales.md` S4.3): `Rem. Anticipos R2!E285`/`K285` **es FORMULA** en la v2 (`E285=E192+E203-K192`), cuando en la medicion de origen era literal. Por eso hoy existe el recorte `ProcesadorPeriodo.cs:198-217` que omite la conciliacion por empresa en Q2 | Plan 50 S8 + T0-Plan50 S4.3 + `Remuneracion.Core/Services/ProcesadorPeriodo.cs:198-217` |
| F5 | **El gate no tiene cobertura de test**: 0 ocurrencias de su mensaje en `Remuneracion.IntegrationTests/TestResults/p53-h2-suite.trx`. Por eso el defecto no aparecia en la linea base (1289/1222/67/0) | trx citado (el T0 lo re-verifica por grep) |
| F6 | El Plan 54 **no causo** estas deltas: agregado por empresa y mapas R2/R4 intactos (Plan 54 S8 G-PROC); con la semantica v1 de ASE1 el Δ era de orden miles de millones (Plan 54 S7 R-CIERRE-EMPRESAS-2026101). El Plan 54 redujo la brecha, no la creo | Plan 54 S7/S8 |
| F7 | Doctrina vigente que NO se re-litiga: `.opencode/project-context.md` (Plan 50 S0.2: fuentes de recaudo opcionales, ausente → 0, nunca abortar; invariantes de cierre del reporte siguen fail-fast) + Plan 53 (mapeo por seccion por operando) + Plan 54 (lectura R1 unica) | citas en cada documento |

### 0.4 Glosario de la plantilla estandar (doctrina de la propietaria, ratificada por el Ingeniero 2026-10-23 — ENMIENDA)

1. **Componente** = encabezado de columna en `Reporte Componentes R1` (TDF, TTL, TVIAT, Aprovechamiento, CCSA Prest.Aprov., SERVICIO ESPECIALES, …). **Los valores de las celdas pertenecen a componentes.**
2. **Fuente de recaudo** = columna **C** (ENEL, OCCIDENTE, NUEVO ESQUEMA, ENERBIT, …).
3. **Canal de recaudo** = columna **D** (H, O, T, E, 1, 2, 5, …).
4. **Concepto** = columna **E** (Vlr Servicio, Vlr Intereses, Total).
5. Cada **cuadro** (seccion) declara **su propia** combinacion de fuentes, canales y conceptos; **no todos los cuadros tienen lo mismo**. Ejemplo de la propietaria: OCCIDENTE, primer cuadro de Promoambiental (filas 9–15), trae canales **H, O, T** y cada canal sus conceptos `Vlr Servicio` / `Total`.
6. **La plantilla es la autoridad; prohibido lo quemado por ASE.** Los destinos se resuelven por las etiquetas de la plantilla. **Prohibido** cualquier "el ASE tal debe traer esto", cualquier lista por ASE, cualquier tabla congelada por (empresa, ASE) y cualquier rama por quincena.

### 0.5 Invalidacion del mecanismo anterior (evidencia del H0 ya ejecutado — ENMIENDA, no re-ejecutar)

El H0 (`Docs/T0-Plan55-Conciliacion-Por-Empresa.md`, fixture 81/81 verde, `p55-H0-suite.trx` 1455/1388/67/0) probo por bytes:

1. **Ninguna de las 112 celdas destino de `WorkbookLeafCellMapPorEmpresa` es termino de la `<f>` del visible de su propio ASE** (teorema `T0_1_NingunDestinoEsTerminoDelVisibleDeBloque`, 15/15 verde); **34/56 de R1 caen en bloques ajenos**; `F353` (destino de ENEL ASE4) es termino del TOT_OPT de **ASE3** (`F415 = F330+F353+F299-L299-L330`, contaminacion cruzada). ⇒ La tabla congelada por (empresa, ASE) **no puede ser el mecanismo, ni parchada con tres entradas** (D4/H1 originales quedan invalidados).
2. Las 56 celdas R1 **ni siquiera se escriben** en la corrida 5-ASE (`OpenXmlPlantillaWriter.cs:266` `espejoR1` → `:323` `omitirR1: true` → `:1250`): las deltas R1 son defecto de **descomposicion en dominio**, no de direccion de escritura.
3. Las cinco deltas (D1..D5) quedan como **sintomas** de esa tabla, no como lista de parches. Clasificacion ratificada por el Ingeniero: **las cinco son causa (i) — nuestras**. `D3` **no** es causa (ii): los bytes prueban `Rem. Anticipos R2!B70="EAB"`, `C70="Total"`, `E70..K70` literales `<v>0</v>` sin `<f>` (slot declarado por el cuadro, sin entrada en el mapa); `Reversion Pagos R4!A8="ENEL"`/`D8` literal y `A110="OCCIDENTE"`/`D110` literal. **No hay divergencia de datos que declarar** (la marca "D3 (ii) pendiente de ratificacion" del H0 queda vaciada por este acto).
4. Siguen vigentes del T0: 6 destinos que son formula en la v2 (`R1!F86/F95/L95/F416`, `R2!E285/K285`), `R-RECORTE-Q2` sin levantar (T0.5) y `R-SIGMA-Q2`.

---

## 1. PROPOSE

### 1.1 Intent

Que la conciliacion por empresa **cierre contra el bloque en Q1 para la corrida 5-ASE de `2026101`**: cada una de las 5 deltas (D1-D5) queda atribuida con evidencia de bytes a exactamente una causa, el mapa por empresa resultante queda re-congelado contra la v2 vigente, el gate `ValidarSigmaEmpresas` pasa a tener cobertura de test por astilla, y el `DetRetri` final no se mueve ni un peso. **T0 por bytes primero, cero codigo de produccion en H0.**

### 1.2 In Scope (T0 + T1 + T2 + T3, prohibido ampliar sin nuevo OK del Ingeniero)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **T0** | Re-congelamiento por bytes (ZIP+XML, sin usar `<v>` de celdas con `<f>` como oraculo) de la celda destino de cada operando por empresa × ASE × componente (R1/R2/R4): veredicto `<f>` vs literal, texto de formula, y si la celda es la misma que consume el visible de bloque | Tabla celda-por-celda con veredicto + evidencia; puerta H0→H1 |
| **T1** | Clasificacion de cada una de las 5 deltas en exactamente una causa: (i) direccion equivocada — el valor por empresa cae donde el total no lo lee; (ii) la fuente no cuadra — S real de empresas ≠ total del bloque en la fuente; (iii) celda destino es formula / guard. Nada de "ajustar hasta que cuadre". Delta no atribuible a una sola causa ⇒ **STOP** | 5 veredictos con derivacion (fuente × `<f>` v2 × mapa) |
| **T2** | Explicacion de los dos delatores: ASE4-R1 = 309.350,00 = `ReciprocidadEaab` exacto (¿valor huerfano? ¿duplicado? ¿bloque que no lo recorre?) y ASE1-R1 = 2.015,07 = \|primera fila `Mes`\| (¿la seccion "primera" entra en el agregado por empresa y no en el bloque, o al reves?) | Dos explicaciones con bytes que el T0 demuestra |
| **T3** | Mapeo correcto resultante + criterio de cierre + plan de implementacion (H1) y de cobertura de test (H2) para el gate, que hoy no tiene ninguno | Mapa + S2/S3 verificables |
| **T4** | Veredicto documentado (sin tocar) sobre `R-RECORTE-Q2`: si el T0 demuestra que el mapa re-congelado ya escribe en celdas no-formula, evaluar levantar el recorte como parte del mismo plan **solo si el T0 lo respalda** — si no, se declara con su gate | Residuo con evidencia y gate de reactivacion |

> **Nota de enmienda 2026-10-23 (H0 ejecutado, no se re-ejecuta):** T0–T2 cerrados por el H0 con las cinco deltas en causa **(i)** (D3 paso de "(ii) pendiente" a (i) por acto del Ingeniero, S0.5). T3 ya no es "mapa re-congelado" sino **resolvedor por etiquetas** (S2-enmendado). T4 confirmado por T0.5 (recorte vivo).

### 1.3 Non-goals / out (vinculante)

- No tocar la plantilla (SHAs v1/v2 antes/despues; cero `<f>`, cero `InsertRow/DeleteRow`).
- **Prohibido tocar** `EscritorValoresSlotR1`, `ContratoSeccionPeriodFreeR1`, `PreparadorBasePipeline`, `AnclasExtempR1`.
- Prohibido ajustar valores, tolerancias o mapas "para que cuadre". Toda delta debe quedar atribuida con evidencia de bytes.
- No corregir EXTEMP/R2 ni reclasificar divergencias ajenas (Planes 49/50/53 y sus deudas intactas).
- No reabrir la semantica de agregados (Plan 25 D-C/D-F) ni el contrato por operando (Plan 53) ni la lectura unica R1 (Plan 54).
- Commits, UI, CLI, docs de usuario. NUNCA `git commit` / `git push`.

### 1.4 Riesgos con mitigacion

| ID | Riesgo | Mitigacion |
|---|---|---|
| R-AJUSTE | "Mover el mapa hasta que cuadre" sin causa (el anti-patron que este plan prohibe) | T1 exige una sola causa con derivacion `fuente × <f> × mapa`; delta no atribuible ⇒ STOP; CA0 (todo numero con derivacion) |
| R-ORACULO | Usar `<v>` stale de celdas con `<f>` como oraculo (leccion Planes 28/29: prohibido) | Metodo T0: texto-`<f>` + literales + fuente leida; `<v>` de celda con `<f>` nunca es oraculo (S0 G-WORKBOOK) |
| R-FALSO-VERDE | Levantar `R-RECORTE-Q2` sin evidencia y reintroducir escrituras en celdas-formula (clase `R-MAPA-V2-R2-ASE4`) | T4 condicionado: solo si el T0 demuestra destinos no-formula; si no, residuo declarado con gate |
| R-SIN-COBERTURA | El gate sigue sin tests y la regresion es silenciosa (F5: 0 ocurrencias en el trx) | H2 exige al menos un test por astilla del gate (S3); sin H2 verde no hay H3 |
| R-DETRI | El fix del mapa mueve el `DetRetri` (el dinero final) | CA-Derivado: `DetRetri` final = R10 oficial en los 5 ASE, ni un peso de movimiento; mismatch ⇒ STOP |
| R-PLANTILLA | El re-congelamiento toca una celda-formula o la geometria | SHAs + G-FORMULA + guard `ERR-PLANTILLA` intacto; H0 es read-only (cero codigo de produccion) |
| R-RECAIDA | Un fix futuro reintroduce direcciones sin evidencia | Mapa re-congelado con cita triple por celda (geometria + `<f>` consumidora + manual/oraculo, patron Plan 53 D2) + tests que pinchan cada delta |

---

## 2. DESIGN

### D1. Metodo T0 por bytes (EJECUTADO en H0 — se conserva integro como evidencia)

> H0 CERRADO 2026-10-23: fixture `ConciliacionEmpresaCierreBloqueT0Tests` 81/81 verde, evidencia `Docs/T0-Plan55-Conciliacion-Por-Empresa.md`, suite `p55-H0-suite.trx` (1455/1388/67/0). Lo que sigue es el metodo ejecutado, no se re-ejecuta.

**Forma normativa:** BCL puro (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`), sin Excel/COM, sin OpenXml SDK, sin ExcelDataReader. Rutas explicitas (prohibido `Insumos.Raiz()` en lo nuevo). Cero escritura sobre fuentes y plantillas. Cero codigo de produccion en H0: solo fixture de T0 + documento de evidencia `Docs/T0-Plan55-Conciliacion-Empresa-Cierre-Bloque.md`.

Por cada (ASE × componente R1/R2/R4 × empresa con operando declarado en `EditablesR*PorEmpresa`):

1. **Destino en la v2 vigente** (SHA `0F192AD2…7F2DAB`): hoja + celda, veredicto `<f>` vs literal, texto exacto de `<f>` si la hay, y si la celda es termino de la `<f>` del visible de bloque (TOT_OPT/EXTEMP/R2-total/R4-total) o cae fuera de su lectura.
2. **Origen en la fuente octubre-1**: filas-etiqueta del bloque de cierre (R1: col C + `D=Total`; R2: col B + `C=Total`; R3/R4: col A + `B=Total` — criterio espejo del reader, T0-Plan50 S1), valores F/Especiales, fila `Total` del bloque y su valor.
3. **Derivacion de la delta**: `S empresas (dominio actual)` vs `visible de bloque (dominio)` vs `S etiquetas-fuente (bytes)` vs `fila Total-fuente (bytes)` vs `evaluacion de la <f> del bloque (texto-<f> + literales, nunca <v> stale)`.

| Candidata | Veredicto | Motivo |
|---|---|---|
| **M-A: T0 bytes celda-por-celda + clasificacion causal cerrada (i/ii/iii)** | **Elegida** | Es el unico metodo que distingue "direccion equivocada" de "fuente que no cuadra" sin ajustar nada; precedente: T0-Plan50, T0-Plan52, Plan 53 D2 |
| M-B: re-derivar el gate o relajar la tolerancia | Descartada | Prohibido por restriccion dura; convertiria un defecto de mapa en un falso verde |
| M-C: reescribir mapas por empresa "a ciegas" | Descartada | G2-D1/Plan 50 S1.3: nunca parche sin T0; es la clase de defecto que nos trajo aqui |

### D2. Taxonomia causal: veredicto final (ENMIENDA 2026-10-23 — las cinco son (i))

Cada delta D1-D5 quedo clasificada por el H0 con derivacion completa (`Docs/T0-Plan55-Conciliacion-Por-Empresa.md` S5); el Ingeniero ratifico: **las cinco son causa (i) — nuestras**. La marca "D3 (ii) pendiente de ratificacion" queda **vaciada**: los bytes prueban slot declarado sin entrada en el mapa (`R2!B70="EAB"`/`C70="Total"`/`E70..K70` literales 0 sin `<f>`), que es cobertura del mecanismo, no divergencia de datos (S0.5.3).

- **(i) Mecanismo nuestro** — D1 (seccion primera entra al bloque `F16` y no al agregado: cadena de 3 terminos + mapa `(2,1)` sin `TotalMenor`); D2 (`UltimaOportuna` absorbe `Aplicacion#1` que el TOT_OPT no recorre); D3 (cuadro declara `EAB`, sin cruce que lo resuelva); D4/D5 (slots `R4!D8`/`D110` existentes, sin entrada `(2,1)`/`(4,2)` en el mapa).
- **(ii) La fuente no cuadra** — cajon vacio en este plan: ninguna delta lo ocupa. Si un periodo futuro lo ocupara, exige veredicto de negocio del Ingeniero (condicion S6 original, vigente).
- **(iii) Celda destino es formula / guard** — no ocupa ninguna delta; ocupa 6 destinos (`R1!F86/F95/L95/F416`, `R2!E285/K285`) que condicionan H1 (bloqueador T0.5-S7.1: sin re-congelar `E285`/`K285` la corrida 5-ASE aborta en `ERR-PLANTILLA` aunque cierren las 5 deltas).

**Regla de cierre (vigente):** delta no atribuible a una sola causa con derivacion completa ⇒ **STOP** + escalado al Ingeniero (no se implementa H1).

### D3. Los dos delatores: RESUELTOS por el H0 (no son hipotesis)

- **H-ASE4-R1 (309.350,00):** identidad confirmada (`356.584,67 + (−47.234,67)`, exacto), lectura "huerfano/duplicado" **refutada**: el dinero existe una sola vez en `Aplicacion#1`; `NUEVO ESQUEMA` no tiene filas en `Oportuno`; el TOT_OPT no recorre Aplicacion y el operando (`UltimaOportuna` sobre el bloque de la empresa) si lo suma. Causa **(i)** (T0 S6.1).
- **H-ASE1-R1 (2.015,07):** identidad confirmada (|`Sheet1!F8`|), direccion **al reves de la primera alternativa**: la primera (`Oportuno#1`, fila 8) **SI entra en el bloque** (termino `F16` de `F75`) y **NO entra en el agregado** (cadena de 3 terminos ignora `totalMenor`; mapa `(2,1)` sin `TotalMenor`). Causa **(i)** de seccion (T0 S6.2).

### D4. Resolucion por etiqueta, camino por empresa (ENMIENDA 2026-10-23 — sustituye al "mapa corregido aditivo", INVALIDADO en S0.5)

**Forma normativa:** se sustituye la resolucion por tabla congelada por un **resolvedor por cruce de etiquetas**:

`(cuadro/seccion, fuente C, canal D, concepto E) → celda`, con el **componente** resuelto por **encabezado de columna** del template.

**Maquinaria a reusar (no reinventar):** `ColumnasComponentes` F..Z por encabezado + `SinonimosEmpresaR1` (etiquetas) + `ContratoSeccionPeriodFreeR1.SeccionPorCierre` (seccion) + `CanalEfectivo` (canal) — el mismo criterio que ya usa `EscritorValoresSlotR1` para el camino de bloque. El camino por empresa pasa a ser un segundo consumidor de esa maquinaria, no un mapa paralelo.

| Candidata | Veredicto | Motivo |
|---|---|---|
| **R-A: resolvedor por cruce de etiquetas, reusando la maquinaria del camino de bloque** | **Elegida** | La plantilla es la autoridad (S0.4): el cuadro declara sus etiquetas y el cruce las resuelve sin listas por ASE. Cero quemado; OCP por fila declarada en el template, no en codigo |
| R-B: parchar `EditablesR*PorEmpresa` con tres entradas | **Invalidada** | S0.5.1: ninguna de las 112 celdas es termino del visible propio; el mecanismo no puede cerrar por construccion |
| R-C: tabla nueva por (empresa, ASE) "corregida" | Descartada | Es R-B con otro nombre; viola la prohibicion expresa de la propietaria |

**Regla de ausencia (ratificada, Plan 50 S0.2):** `(fuente, canal, concepto)` declarado por el cuadro y no presente en la fuente → **cero del template**, no error. Nunca abortar por ausencia de fuente/canal/seccion; los invariantes de cierre del reporte siguen fail-fast.

**Regla de exceso:** la fuente trae `(fuente, canal, concepto)` que el cuadro **no** declara → **registro auditable + divergencia declarada** (analizar contra Plan 44 `EncabezadosSinDestino` y Plan 42: catalogo con reconocimiento explicito, dueno y condicion de retiro), **nunca** descarte silencioso ni celda inventada ni aborto por defecto.

**Agregado por empresa (definicion anti-tautologia):** se calcula **desde la fuente**, agrupando por etiqueta de fuente, sobre **el mismo conjunto de roles/secciones que alimenta el visible del bloque** (resuelto por firma/seccion, no por lista de empresas). **Prohibido** definirlo como "la suma de lo que escribimos en los terminos de la formula": eso haria el gate trivialmente verde (falso verde) — el gate debe seguir siendo una verificacion independiente del camino de escritura.

Criterio de cierre por delta: `S empresas == visible de bloque ±0,5` **con la derivacion que lo explica** (no basta el verde). Criterio global: cero `ERR-VALIDACION` de `ValidarSigmaEmpresas` en la corrida 5-ASE de `2026101` + `DetRetri` = R10 oficial en los 5 ASE (S3-CA).

### D5. Cobertura del gate + candado anti-quemado (nuevo T4, ENMIENDA)

El gate hoy no tiene tests (F5; T0.6 lo re-verifico: 0 ocurrencias en `p53-h2-suite.trx`). H2 anade **al menos un test por astilla**: R1-suma-vs-bloque, R2-suma-vs-bloque, R4-suma-vs-bloque, tolerancia ±0,5 en el borde, early-return Q2 (`NumeroQuincena != 1 ⇒ return`), early-return lista vacia, y **un test por delta D1-D5** que reproduce la firma (S empresas vs bloque vs fuente) y pincha la causa (i). Patron: fixture chips con `ConciliacionEmpresaInputs` + `WorkbookLeafInputs` construidos a mano (sin workbooks), mas e2e `2026101` 5-ASE como oraculo final. Prohibido `Skip`/`Assert.True(true)`/borrados.

**Candado anti-quemado (nuevo T4):** test estatico que **falla si en el camino de resolucion aparecen listas por ASE, literales de empresa, o mapas congelados por (empresa, ASE)**. Mismo patron del candado del Plan 54 (`LecturaR1SinRamaQuincenaTests`): allow-list cerrada (la maquinaria reusada de D4: `SinonimosEmpresaR1`, `SeccionPorCierre`/`CanalEfectivo`, columnas por encabezado) + prueba de injerto (una lista por ASE o un literal de empresa injertado a proposito debe enrojecerlo). Sin candado verde no hay H3.

### D6. `R-RECORTE-Q2` (CONFIRMADO vivo por T0.5 — no se levanta)

El recorte `ProcesadorPeriodo.cs:198-217` (conciliacion por empresa omitida en Q2) **sigue vigente**: el T0.5 demostro que no hay respaldo (`R2!E285`/`K285` formula en la v2 vigente) ⇒ **no se levanta en este plan** y `ProcesadorPeriodo.cs:198-217` **no se toca**. El gate Sigma **no** se reactiva en Q2 — `R-SIGMA-Q2` (T0-Plan50 S4.2) sigue vigente: no es descomposicion valida del visible alli. Gate de reactivacion futuro: re-congelar `E285`/`K285` (H1.4) + plan propio para Q2.

### D7. Lo que NO se toca (lista cerrada, verificable por grep de cierre)

`EscritorValoresSlotR1`, `ContratoSeccionPeriodFreeR1` (`SeccionPorCierre`/`Alinear`), `PreparadorBasePipeline`, `AnclasExtempR1`. Tolerancia ±0,5 intacta. `<f>` de la plantilla intactas. `R1Q2EditablesPorAse` y lectura unica R1 (Plan 54) intactos salvo lo que el T0 ordene citar (por defecto, nada). **Cambio autorizado por la enmienda (no es "tocar", es retirar del camino):** `WorkbookLeafCellMapPorEmpresa` sale del camino productivo y se marca LEGACY (como se hizo con `WorkbookLeafCellMapPorAse` en el Plan 54), conservandose solo para los tests historicos que lo referencien; ningun codigo nuevo lo consume.

### Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Reader = evidencia por firma/seccion (resolvedor por etiquetas, camino por empresa reunificado con el de bloque); Validador = juicio (intacto, solo gana cobertura); Mapa congelado retirado a LEGACY; Writer = values-only (intacto) |
| **OCP** | ✅ Forma-fuente/cuadro nuevo absorbido por el cruce de etiquetas (el cuadro declara, el resolvedor resuelve), nunca por rama ni por entrada quemada; celda nueva = etiqueta del template, nunca `if (empresa/periodo/ase)` |
| **LSP/DIP** | ✅ Contratos `IValidador`/`IPlantillaWriter`/`IWorkbookLeafInputReader` intactos; cambios detras de las mismas abstracciones |
| **Best practices** | ✅ Contract-first (H0 ejecutado como evidencia + T4/H1 con veredicto previo); fail-fast con ASE+reporte+celda; tolerancia ±0,5 solo en comparacion; values-only; D-H (cero literales/ramas/listas por periodo, empresa o ASE en el camino de resolucion); CRLF; sin commits; prohibido `<v>` de celda con `<f>` como oraculo |
| **Performance** | ✅ T0 acotado BCL ZIP+XML + fixtures existentes; regens a temp fresca; cero workbooks extra en runtime |

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda H)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa trx + **set-diff vs `p53-h2-suite.trx` (1289/1222/67/0)** y vs `p54-H3-suite.trx` (1374/1307/67/0): altas/bajas una por una con causa citada; **rojo nuevo fuera de la lista T0 ⇒ STOP**. 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados.
  - **Lectura ratificada en el cierre (H3.2, O3, 2026-10-23):** "0 tests borrados" se lee como **0 bajas netas de cobertura**. Causa: la enmienda de doctrina cambia el MECANISMO de resolución, y los fixtures que citaban celdas de la tabla congelada se re-apuntan a las celdas que el cuadro declara (misma cobertura, mismas magnitudes) y tres métodos se **renombran** para decir lo que ahora prueban (O2). La verificación es el set-diff de resultados (0 nuevos, 0 desaparecidos) + el conteo total (ninguna alta se pierde).
- G-SHA: v1 `3086AC31…C5214` + v2 `0F192AD2…7F2DAB` intactos antes/despues de cada H.
- G-FORMULA: `<f>` byte-identica en todo el diff (cero `new CellFormula`, cero `InsertRow`/`DeleteRow` fuera de Phase 0).
- G-D-H: cero literales/ramas/listas por periodo, empresa o ASE en el camino de resolucion (grep de cierre + candado T4).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale) o la fuente ya leida; **cero workbooks extra en runtime**.
- G-TRAZA: cada error nuevo nombra ASE + reporte + fila/rol + celda donde aplique.
- G-PROC: cero commits; CRLF sin mezclar en lo tocado.
- G-H0-READONLY: H0 no contiene codigo de produccion (solo fixture T0 + doc de evidencia); el diff de `Remuneracion.Core/` + `Remuneracion.Infrastructure/` en H0 es vacio salvo tests `*T0Plan55*`.

### CA — Criterios de aceptacion (vinculantes; con derivacion obligatoria, CA0 del Plan 54)

- CA0: **todo numero de criterio viaja con su derivacion** (`fuente × <f> v2 × mapa`). Un numero sin derivacion no se acepta como criterio.
- CA1: cero `ERR-VALIDACION` de `ValidarSigmaEmpresas` en la corrida 5-ASE de `2026101` (los 5 items del S0.2 desaparecen por causa clasificada, no por tolerancia).
- CA2: `DetRetri` final = R10 oficial en los 5 ASE, al peso — **no debe moverse ni un peso**: `14.842.425.136` / `11.302.637.605` / `8.823.787.300` / `10.483.172.762` / `7.550.668.665` (Plan 54 S8; derivacion Plan 54 S3-CA).
- CA3: cada astilla del gate con al menos un test que la pinche (el gate hoy no tiene cobertura — F5): R1/R2/R4-suma, tolerancia en el borde, early-return Q2 y lista vacia, mas un test por delta D1-D5.
- CA4: SHAs v1/v2 intactos y las 10 `<f>` de R1 intactas (Plan 54 S2.6); verificacion de que **ninguna celda destino del mapa por empresa es formula en la v2** (o lista cerrada de las que si, con veredicto).
- CA5: build con 0 warnings.
- CA6: `R-RECORTE-Q2` con veredicto documentado: levantado solo si el T0 lo respalda (D6) o declarado vivo con su gate.
- CA7 (ENMIENDA): **cero listas por ASE y cero literales de empresa en el camino de resolucion (candado T4 verde)**. `WorkbookLeafCellMapPorEmpresa` fuera del camino productivo (LEGACY); ningun codigo nuevo lo consume.
- CA8 (ENMIENDA): **el gate no puede ser tautologico: su agregado se deriva de la fuente** (mismo conjunto de roles/secciones del bloque, D4), **no de lo escrito**. Test que pincha la tautologia: agregado calculado sobre celdas escritas debe fallar el gate por construccion o estar prohibido por diseno.

### S1 — Criterios H0/T0 (read-only, cero codigo de produccion)

- S1.1: tabla destino v2 por bytes: por cada (ASE × R1/R2/R4 × empresa) — hoja+celda, veredicto `<f>`/literal, texto-`<f>` exacto, y si la celda es termino de la `<f>` del visible de bloque. Re-verificacion de `E285`/`K285` (`E285=E192+E203-K192`) y barrido de gemelas por ASE.
- S1.2: censo octubre-1 por bytes: filas-etiqueta del bloque de cierre + fila `Total` por (ASE × R1/R2/R4), con valores; congelado como fixture.
- S1.3: clasificacion D1-D5, una causa cada una (i/ii/iii) con derivacion completa; delta no atribuible ⇒ STOP.
- S1.4: explicaciones H-ASE4-R1 y H-ASE1-R1 confirmadas o refutadas por bytes (D3).
- S1.5: veredicto T4 (recorte Q2) con evidencia: destinos-formula cerrados o levantamiento respaldado.
- S1.6: lista cerrada de Δ de suite esperados (por defecto: ninguno — el T0 es read-only; cualquier rojo nuevo en H0 ⇒ STOP). **Sin T0 verde no se abre H1.**

### S2 — Criterios H1 (implementacion; requiere re-ratificacion de la enmienda)

- S2.1: resolvedor por etiquetas para el camino por empresa (R1/R2/R4) reusando `SinonimosEmpresaR1` + `SeccionPorCierre`/`CanalEfectivo` + columnas por encabezado (D4); regla de ausencia → 0, regla de exceso → registro + divergencia declarada (Plan 44/Plan 42).
- S2.2: `WorkbookLeafCellMapPorEmpresa` retirado del camino productivo (LEGACY, solo tests historicos); CA7 verde (candado T4).
- S2.3: agregado por empresa derivado de la fuente con definicion anti-tautologia (D4); CA8 verde.
- S2.4: re-congelados por bytes los 6 destinos formula alcanzables (`R1!F86/F95/L95/F416`; `R2!E285/K285` solo condicionados a `R-RECORTE-Q2`, que sigue vivo); CA4 actualizado: ningun destino del camino por empresa es formula escribible sin gate.
- S2.5: CA1 + CA2 reproducidos en corrida real `2026101` 5-ASE (SHAs + G-FORMULA + 0 warnings).
- S2.6: archivos prohibidos intactos (D7, grep de cierre); tolerancia y mensajes del gate intactos en formato.

### S3 — Criterios H2 (regresion + cobertura)

- S3.1: CA3 verde (cobertura por astilla + por delta) + CA7/CA8 verdes (candado anti-quemado + anti-tautologia).
- S3.2: set-diff == lista T0; `2026071`/`2026072` verdes por la misma ruta (no-regresion Plan 54 CA4); oraculo v2 julio-2 intacto.
- S3.3: code-review independiente **APROBADO** (0 bloqueantes): spec compliance + calidad + greps D-H + SHAs.

### S4 — Criterios H3 (docs/cierre)

- S4.1: doctrina en `.opencode/project-context.md` (solo lo probado: resolvedor por etiquetas, causas D1-D5 en (i), cobertura del gate + candado anti-quemado, veredicto R-RECORTE-Q2).
- S4.2: este plan como fuente de verdad; desviaciones con fecha/motivo/evidencia; cero codigo de produccion en H3.

---

## 4. TASKS (fases H0..H3 con puertas)

### H0/F0 — T0 read-only (EJECUTADO 2026-10-23 — CERRADO, se conserva como evidencia, no se re-ejecuta)

> Resultado H0: fixture 81/81 verde, 0 skip; suite `p55-H0-suite.trx` (1455/1388/67/0, mismos 67 rojos); SHAs intactos; las 10 `<f>` R1 intactas. Las tareas H0 se conservan abajo como registro de lo ejecutado.

Fixture: `Remuneracion.IntegrationTests/ConciliacionEmpresaCierreBloqueT0Tests.cs` (BCL ZIP+XML, rutas explicitas; prohibido `Insumos.Raiz()` en lo nuevo). Evidencia: `Docs/T0-Plan55-Conciliacion-Por-Empresa.md` (correccion de nombre por enmienda: el plan decia `Docs/T0-Plan55-Conciliacion-Empresa-Cierre-Bloque.md`).

| ID | Tarea | Verifica |
|---|---|---|
| T0.1 | Destinos v2 por bytes: por cada (ASE × R1/R2/R4 × empresa) hoja+celda, `<f>` vs literal, texto-`<f>`, y si la celda es termino del visible de bloque; re-verificar `E285`/`K285` y gemelas | S1.1, CA4 |
| T0.2 | Censo octubre-1 por bytes: filas-etiqueta + fila `Total` por (ASE × R1/R2/R4) con valores; congelar como fixture | S1.2 |
| T0.3 | Clasificar D1-D5 en (i/ii/iii) con derivacion `fuente × <f> × mapa`; delta no atribuible ⇒ STOP + escalado | S1.3 |
| T0.4 | Resolver H-ASE4-R1 (309.350 huerfano/duplicado/no-recorrido) y H-ASE1-R1 (2.015,07 primera-`Mes`) por bytes | S1.4 |
| T0.5 | Veredicto T4: barrido formula/no-formula de destinos + recomendacion recorte-Q2 (levantar solo con respaldo o declarar) | S1.5 |
| T0.6 | Grep cobertura: 0 ocurrencias del mensaje del gate en `p53-h2-suite.trx` (F5) + lista cerrada de Δ de suite | S1.6 |
| T0.7 | SHAs v1/v2 antes/despues de H0 + las 10 `<f>` R1 intactas | S0 G-SHA/G-FORMULA |

- **Puerta H0→H1:** fixture verde + S1.1-S1.6 cerrados + APROBADO del Ingeniero. Si T0.3 dispara STOP, no se abre H1.

### H1/F1 — Resolvedor por etiquetas + cierre (REESCRITA por enmienda 2026-10-23; se borran las tareas que agregaban entradas por ASE)

| ID | Tarea | Traza |
|---|---|---|
| H1.1 | Resolvedor por etiquetas para el camino por empresa (R1/R2/R4): cruce `(cuadro/seccion, fuente C, canal D, concepto E) → celda` + componente por encabezado, reusando `SinonimosEmpresaR1` + `SeccionPorCierre`/`CanalEfectivo` + `ColumnasComponentes` (D4). Regla de ausencia → 0; regla de exceso → registro + divergencia declarada (Plan 44/Plan 42) | D4, S2.1 |
| H1.2 | Retirar `WorkbookLeafCellMapPorEmpresa` del camino productivo (LEGACY como `WorkbookLeafCellMapPorAse` en Plan 54; solo tests historicos); cero listas por ASE y cero literales de empresa en el camino (D7, CA7) | D7, S2.2 |
| H1.3 | Agregado por empresa derivado de la fuente sobre el mismo conjunto de roles/secciones del bloque (definicion anti-tautologia, D4); CA8 verde | D4, S2.3 |
| H1.4 | Re-congelar por bytes los 6 destinos formula alcanzables (`R1!F86/F95/L95/F416`; `R2!E285/K285` solo condicionados a `R-RECORTE-Q2`, que sigue vivo — sin tocar `ProcesadorPeriodo.cs:198-217`) | T0 S3.3/S7.1, S2.4 |
| H1.5 | e2e `2026101` 5-ASE: CA1 + CA2 (cero ERR-VALIDACION, DetRetri = R10 al peso) | S2.5 |

- **Puerta H1→H2:** S2 verde + CA1/CA2/CA4/CA7/CA8 + 0w + SHAs + G-FORMULA + greps D-H + candado T4 en construccion.

### H2/F2 — Regresion, cobertura y evidencia

| ID | Tarea | Verifica |
|---|---|---|
| H2.1 | Tests por astilla del gate + por delta D1-D5 (D5) + candado anti-quemado T4 (D5: enrojece ante lista por ASE o literal de empresa injertado; patron Plan 54) | CA3/CA7/CA8, S3.1 |
| H2.2 | Suite trx + set-diff vs `p53-h2-suite.trx` y `p54-H3-suite.trx`; fuera de lista = STOP | S0 G-SUITE |
| H2.3 | No-regresion `2026071`/`2026072` misma ruta + oraculo v2 julio-2 intacto | S3.2 |
| H2.4 | Code-review independiente (`code-reviewer`, este plan como spec): **APROBADO** requerido | S3.3 |

- **Puerta H2→H3:** set-diff exacto + CA1-CA8 + review aprobado.

### H3/F3 — Docs/cierre

| ID | Tarea | Verifica |
|---|---|---|
| H3.1 | Doctrina en `.opencode/project-context.md` (solo lo probado) | S4.1 |
| H3.2 | Cierre de este plan (S8) con desviaciones fechadas si las hubo; cero codigo de produccion | S4.2 |

**Orden sugerido:** H0 → H1 → H2 → H3. Regresion continua (suite + SHAs en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero).

---

## 5. Riesgos (consolidado + gates por fase)

Ver S1.4 (R-AJUSTE, R-ORACULO, R-FALSO-VERDE, R-SIN-COBERTURA, R-DETRI, R-PLANTILLA, R-RECAIDA) + enmienda (R-TAUTOLOGIA, R-QUEMADO). Riesgo residual re-encuadrado por la enmienda: las etiquetas sin dueno en el mapa (`EAB` en R2, `ENERBIT` en R2-ASE5, `ENEL` sin entrada en R4 ASE1, `OCCIDENTE` sin entrada en R4 ASE2 — T0-Plan50 S3-S4) quedaron clasificadas como causa **(i)** de cobertura del mecanismo (slots declarados por el cuadro: `R2!B70`, `R4!A8`/`A110`), no como divergencia de datos; **prohibido** inventar celdas destino fuera del cruce de etiquetas para "cubrirlas" (regla de exceso D4: registro + divergencia declarada Plan 44/Plan 42).

| ID | Riesgo | Mitigacion |
|---|---|---|
| R-TAUTOLOGIA | El agregado por empresa se define como suma de lo escrito y el gate queda trivialmente verde | D4 anti-tautologia (agregado desde la fuente) + CA8 con test que pincha la tautologia |
| R-QUEMADO | El resolvedor reintroduce listas por ASE, literales de empresa o el mapa congelado por la puerta de atras | Nuevo T4: candado estatico allow-list + injerto (patron Plan 54); CA7; G-D-H extendido |

- **Gate H0→H1:** T0 verde + S1.3/S1.4 sin STOP + APROBADO del Ingeniero.
- **Gate H1→H2:** S2 + CA1/CA2/CA4 + 0w + SHAs + G-FORMULA + greps.
- **Gate H2→H3:** set-diff exacto + CA3-CA6 + review APROBADO.
- **Gate H3:** diff de docs citado; cero produccion.

---

## 6. Clarification Gate (requiere aprobacion del Ingeniero; sin esto no se abre H0)

1. **T0 primero:** ¿se ratifica H0 read-only bloqueante (cero codigo de produccion) con la taxonomia cerrada (i/ii/iii) y la regla "delta no atribuible ⇒ STOP"?
2. **D3:** ¿las hipotesis H-ASE4-R1 y H-ASE1-R1 se investigan como hipotesis (a confirmar/refutar por bytes) y no como veredictos?
3. **D6/T4:** ¿el recorte Q2 solo se levanta si el T0 demuestra destinos no-formula, y en ese caso el gate Sigma sigue sin reactivarse en Q2 (`R-SIGMA-Q2` vigente)?
4. **Linea base:** ¿se ratifican `p53-h2-suite.trx` (1289/1222/67/0) y `p54-H3-suite.trx` (1374/1307/67/0, mismo conjunto de rojos) para el set-diff, y que rojo fuera de lista = STOP?
5. **Puerta:** sin 1-4 no se abre H0; sin H0 verde no hay H1; sin review no hay H3.

> **ENMIENDA 2026-10-23 — solo lo nuevo (requiere re-ratificacion; H0 no se re-ejecuta):**
> 6. **Glosario S0.4:** ¿se ratifican los 4 niveles (componente/fuente/canal/concepto), que cada cuadro declara lo suyo, y que "la plantilla es la autoridad; prohibido lo quemado por ASE"?
> 7. **Invalidacion S0.5:** ¿se ratifica que la tabla congelada por (empresa, ASE) no puede ser el mecanismo ni parchada (112 celdas, 34/56 R1 ajenas, `F353` cruzada), que las cinco deltas son (i) —incluida D3— y que no hay divergencia de datos que declarar?
> 8. **D4/D5/H1:** ¿resolvedor por cruce de etiquetas reusando la maquinaria del camino de bloque, regla de ausencia → 0, regla de exceso → registro + divergencia (Plan 44/42), agregado anti-tautologico, retiro del mapa a LEGACY, re-congelamiento de los 6 destinos formula y candado anti-quemado (nuevo T4)?
> 9. **Puerta:** sin 6-8 no se abre la H1 reescrita; sin candado T4 no hay H3.

> **RATIFICADO 2026-10-23 por el Ingeniero (con calidad senior) — gate 1-5:** 1) H0 read-only bloqueante con taxonomía cerrada (i/ii/iii) y regla "delta no atribuible ⇒ STOP". 2) H-ASE4-R1 y H-ASE1-R1 se investigan como hipótesis, no veredictos. 3) El recorte Q2 solo se levanta si el T0 demuestra destinos no-fórmula; `R-SIGMA-Q2` sigue vigente en Q2. 4) Líneas base `p53-h2-suite.trx` y `p54-H3-suite.trx` ratificadas; rojo fuera de lista = STOP. 5) Puertas H0→H1→H2→H3.

> **RATIFICADO 2026-10-23 por el Ingeniero — enmienda de doctrina, gate 6-8:** 6) glosario de los 4 niveles ratificado; **la plantilla es la autoridad y queda prohibido lo quemado por ASE**. 7) invalidación de la tabla congelada ratificada (112 celdas, 34/56 R1 ajenas, `F353` cruzada); **las cinco deltas son (i)** — incluida D3 — y **no hay divergencia de datos que declarar**. 8) resolvedor por cruce de etiquetas + reglas ausencia→0 / exceso→registro+divergencia + agregado **anti-tautológico** + tabla a LEGACY + re-congelamiento de los 6 destinos fórmula + candado anti-quemado (T4). 9) puertas vigentes.

> **Condiciones del Arquitecto (vinculantes):** (a) la clasificación de una delta como causa **(ii)** es veredicto de negocio: el implementer no la declara por su cuenta; (b) el catálogo de sinónimos queda **ABIERTO**: iguala rótulos, **nunca** exige presencia de una fuente ni infiere ausencia — un sinónimo faltante no puede convertirse en validación.

**Gate §6 completo (1-9) ratificado ⇒ Plan 55 APROBADO 2026-10-23. H0 cerrado; H1 reescrita habilitada.**

---

## 7. Residuos declarados (fuera de alcance — NO implementar)

**R-RECORTE-Q2 (vivo, sin tocar):** `Remuneracion.Core/Services/ProcesadorPeriodo.cs:198-217` mantiene `if (!esQuincena2)` que omite la conciliacion por empresa en Q2. Deuda `R-MAPA-V2-R2-ASE4` (Plan 50 S8): `Rem. Anticipos R2!E285`/`K285` es formula en la v2 (`E285=E192+E203-K192`). **Evidencia:** Plan 50 S8 + T0-Plan50 S4.3 (+ re-verificacion T0.1/T0.5 de este plan). **Gate de reactivacion:** D6 de este plan. No se toca salvo H1.4 respaldada por el T0.

**R-SIGMA-Q2 (vivo, no se re-litiga):** el gate Sigma no es descomposicion valida del visible en Q2 (T0-Plan50 S4.2: 14 divergencias medidas). Aunque se levante el recorte, el gate NO se reactiva en Q2. Sin tocar.

**R-ANCLA-EXTEMP (vivo, Plan 54 S7):** `AnclasExtempR1` congela `F682`/`F702` como literales pero la v2 vigente trae `F682 = <f>F660+F629-L629</f>`; `ResolverAnclasExtempR1` aborta con `ERR-PLANTILLA`. Con su propio T0 y plan. No se toca aca (D7).

**R-SLOT-Q1-ENERBIT → R-JERARQUIA-AB (renombrado 2026-10-23; ver plans/56, que lo resuelve):** cerrado el gate Sigma y validado el DetRetri contra el R10, la corrida 5-ASE de `2026101` **termina** en `ERR-SIN-SLOT` de `EscritorValoresSlotR1`: `ASE 5`, fuente R1 (`Recaudoporcomponente`), fila `A|B|C|D|E='||ENERBIT|Total|'`, rotulo `ENERBIT`, canal `Total`, concepto vacio, **`1.ª seccion`, ocurrencia #1**, sin ningun slot con firma + seccion compatibles en el contrato T0 del ASE 5. Es deuda **PRE-EXISTENTE** de la familia E1-6 (Planes 41/46; declarada como deuda de escritura en Plan 54 S7 y en `ContratoUnicoLecturaR1H2E2ETests`), **no agravada** por este plan: `EscritorValoresSlotR1` y `ContratoSeccionPeriodFreeR1` no se tocaron y el bloque espejo es el mismo. **Gate de reactivacion:** plan propio que revise el contrato de seccion Q1 <-> slots del template (sin remap, sin insert, sin colocacion por posicion).

**R-ETIQUETAS-SIN-DESTINO (re-encuadrado por enmienda 2026-10-23, T0-Plan50 S3-S4):** `EAB` (R2), `ENERBIT` (R2 ASE5), `ENEL` sin entrada (R4 ASE1), `OCCIDENTE` sin entrada (R4 ASE2). El H0 los clasifico como causa **(i)** de cobertura del mecanismo (el cuadro declara el slot: `R2!B70="EAB"`, `R4!A8="ENEL"`, `R4!A110="OCCIDENTE"`); **no hay divergencia de datos que declarar** (D3 vaciada de (ii) por acto del Ingeniero). Si el cruce de etiquetas encuentra un exceso real futuro, aplica la regla de exceso D4 (registro + divergencia Plan 44/Plan 42); prohibido inventar celdas destino.

---

## 8. Estado

**PROPUESTO 2026-10-23.** Pendiente aprobacion del Ingeniero (Clarification Gate S6). Sin APROBADO no se abre H0. Padre: Plan 54 (CERRADO). Madre doctrina: Plan 50 S0.2/S8 + Plan 53 + Plan 54 + `.opencode/project-context.md`.

**APROBADO 2026-10-23.** Gate §6 ratificado. H0 habilitado. Condición vinculante: H0 cierra con la tabla de atribución (una causa por delta) y **se detiene**; la clasificación (ii) y la apertura de H1 requieren ratificación del Ingeniero.

**H0 EJECUTADO 2026-10-23** (`Docs/T0-Plan55-Conciliacion-Por-Empresa.md`, fixture 81/81, `p55-H0-suite.trx` 1455/1388/67/0).

**ENMIENDA DE DOCTRINA 2026-10-23 — RATIFICADA (gate 6-8) ⇒ Plan 55 APROBADO.** El mecanismo de tabla congelada por (empresa, ASE) queda invalidado; se implementa el **resolvedor por cruce de etiquetas** (cuadro × fuente C × canal D × concepto E, componente por encabezado), con reglas ausencia→0 / exceso→registro+divergencia, agregado **anti-tautológico** y **candado anti-quemado (T4)**. H1 reescrita habilitada. H0 no se re-ejecuta.

### CERRADO 2026-10-23

**Estado:** **CERRADO** (H0 → H1 → H2 → H3). Code-review independiente: **APROBADO, 0 bloqueantes** (H2.4). H3 no toca código de producción salvo el doc-comment huérfano O1 y la documentación.

**Evidencia de cierre (medida, no heredada):**

| Gate | Medición |
|---|---|
| G-BUILD | `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores** |
| G-SUITE | `p55-H3-suite.trx` = **1484 / 1417 / 67 / 0**; set-diff de rojos vs `p54-H3-suite.trx` y `p53-h2-suite.trx` = **VACÍO** (mismos 67 nombres; 0 nuevos, 0 desaparecidos); +29 altas verdes vs `p55-H0-suite.trx` (1455/1388/67/0) |
| G-SHA | v2 `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB` y v1 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` **intactos** |
| G-FORMULA | cero `<f>` nueva, cero `InsertRow`/`DeleteRow`; toda celda destino resuelta es **literal** (`DestinosFormula` vacío en los 5 ASE) y los 6 destinos fórmula de la v2 (`R1!F86/F95/L95/F416`, `R2!E285/K285`) **no son alcanzados** |
| CA1 | corrida CLI real `2026101 --cinco-ase` → **cero `ERR-VALIDACION` de `ValidarSigmaEmpresas`** |
| CA2 | `DetRetri` = R10 oficial en los 5 ASE, **sin mover un peso**: `14.842.425.136` / `11.302.637.605` / `8.823.787.300` / `10.483.172.762` / `7.550.668.665` (validado además contra `R10_Remuneracion_2026101.xlsx`) |
| CA3/CA7/CA8 | `ConciliacionEmpresaSigmaPlan55Tests` (astillas del gate + D1-D5 + anti-tautología + H1.4) y `ConciliacionEmpresaCandadoAntiQuemadoTests` (candado + 3 injertos) **verdes** |
| D1-D5 | las cinco deltas **cerradas al peso** con derivación `fuente × <f> v2 × celda` |

**Cierre de las cinco deltas (derivación CA0):**

| Δ | ASE·reporte | Valor que cierra | Derivación |
|---|---|---:|---|
| D1 | ASE1·R1 | `14.815.021.974,75` | `ΣF(Mes) − ΣEsp(Mes sin la última)` = `14.834.953.667,78 − 19.931.693,03`; el agregado por empresa ahora incluye la 1.ª sección (`F8 = −2.015,07`) |
| D2 | ASE4·R1 | `10.473.967.813,42` | RECIP (rótulo `NUEVO ESQUEMA`) solo trae sección **Aplicación** ⇒ aporta `0` al visible (el TOT_OPT solo recorre secciones `Oportuno`) |
| D3 | ASE2·R2 | `95.634.974,32` | `Componente + Subs/Cont − Especiales` = `116.410.420,67 − 20.704.709,40 − 70.736,95`; `EAB` (declarada por `R2!B70`) aporta `149.137,87 − 18.613,87 = 130.524,00` |
| D4 | ASE1·R4 | `−3.833.780,00` | `ENEL` resuelve su slot declarado (`R4!A8/B8`) ⇒ `−3.698.870,00`; total `D17` |
| D5 | ASE2·R4 | `−21.803.648,45` | `OCCIDENTE` resuelve su slot declarado (`R4!A110/B110`) ⇒ `−945.063,00`; total `D17` |

**Desviaciones fechadas (2026-10-23):**

- **(a) `RutaPlantilla` como propiedad del contrato, no como parámetro.** Un parámetro obligaba a editar la llamada en `ProcesadorPeriodo.cs:211`, dentro del bloque protegido `198-217`. Se setea UNA vez por corrida en **`ProcesadorPeriodo.cs:80-84`** (fuera del bloque). `IWorkbookLeafInputReader.RutaPlantilla`.
- **(b) H1.4 resuelto por bytes en RUNTIME, sin tabla nueva.** Congelar otra tabla habría reintroducido el mecanismo invalidado por S0.5. El "re-congelamiento" queda garantizado por: geometría leída de la plantilla vigente, guard anti-fórmula por bytes y test que prueba que ningún destino es fórmula y que los 6 fórmula no se tocan.
- **(c) `FuentesAusentes` solo donde el CUADRO declara el rótulo.** Antes se emitía por mapa; una fuente que el cuadro no declara no es "ausente" (no aplica). Se ajustaron `FuentesAusentesPlan50Tests` y `ControlAntiFalsoCeroPlan50Tests` (añaden `RutaPlantilla`; las aserciones de celdas pasan a los destinos resueltos).
- **(d) Exceso registrado (nunca celda inventada).** `R2:EAB:Subs/Cont` (EAB no tiene fila en el cuadro Subs/Cont del ASE2) y `R1:ENERBIT:Componente` (ENERBIT no tiene fila en el cuadro Componente del ASE5): ambas entran al agregado por venir de la fuente y quedan auditadas en `ConciliacionEmpresaInputs.DestinosSinCelda`. Por eso ASE2 devuelve **6** entradas (5 del catálogo + `EAB`).
- **(e) Fixture del Plan 53 actualizado con magnitudes IDÉNTICAS.** `ContratoSeccionPorOperandoR1Tests`: los 4 valores del manual del administrativo se reproducen al peso (`11.701.157.844,90` / `11.748.395.433,49` / `19.237.658.087,83` / `2.217.280.433,95`); solo cambian las celdas destino (las que el cuadro declara) y las aserciones de "Menor/Medio" se sustituyen por el contrato vigente.
- **(f) INCIDENTE y su verificación (declarado).** Durante H1, una edición mal codificada derivó en `git checkout -- ExcelDataReaderWorkbookLeafInputReader.cs`, que borró cambios **sin commitear** del Plan 54 en ese archivo. El archivo se reconstruyó desde `HEAD` + los cambios del Plan 54 y se auditó de forma independiente en el code-review (H2.4): **8/8 anclas del Plan 54/55 presentes**, sin duplicados ni código muerto, compilación 0 warnings y **set-diff de rojos vacío** en la suite completa. **Aprendizaje registrado:** prohibido `git checkout`/`git reset` en este repo (trabaja sin commits) sin copia previa del archivo.

**Observaciones del code-review (H2.4) y su cierre:**

- **(O1)** Doc-comment huérfano de `ExcelDataReaderWorkbookLeafInputReader.cs` (citaba `ExtraerCeldasR1`, método retirado): **corregido en H3**, ahora nombra al consumidor real (`ResolvedorConciliacionEmpresa`) y deja constancia del retiro.
- **(O2)** Tres métodos de `ContratoSeccionPorOperandoR1Tests` **renombrados** (misma cobertura, no baja): `Oct1_Ase5_LasTresEmpresas_MainUnicaRecaudo` → `Oct1_Ase5_LasTresEmpresas_ResuelvenSuFilaDeclarada`; `Jul2_Ase5_Enel_MenorCeroPorSeccionUnica` → `Jul2_Ase5_Enel_SeccionUnicaSinDuplicarElBloque`; `ContratoMapa_DeclaraSeccionPorCelda_EnLas13Claves` → `ContratoMapaLegacy_DeclaraSeccionPorCelda_EnLas13Claves`.
- **(O3)** G-SUITE se lee como **"0 bajas netas de cobertura"** (no "0 tests borrados"), con la causa registrada en §S0: el cambio de mecanismo re-apunta aserciones y renombra métodos, y la garantía pasa a ser el set-diff de resultados + el conteo total.

**Frase de cierre:** *El Plan 55 queda CERRADO el 2026-10-23: la conciliación por empresa se resuelve por etiquetas de la plantilla (la plantilla es la autoridad), las cinco deltas D1-D5 cierran al peso con derivación por bytes, el gate Σ empresas deja de ser tautológico y pasa a tener cobertura por astilla con candado anti-quemado, el `DetRetri` sigue siendo el R10 oficial en los 5 ASE sin mover un peso, y el bloqueo que queda —`R-SLOT-Q1-ENERBIT`— es deuda pre-existente de escritura, declarada con su propio gate.*
