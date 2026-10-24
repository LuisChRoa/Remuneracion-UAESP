---
titulo: Plan 57 - Anclas de validacion derivadas de la plantilla (fin de la tabla congelada)
estado: APROBADO 2026-10-23 (gate §6 ratificado; H0 habilitado, H1 condicionada al T0)
fecha: 2026-10-23
numero: 57
planes_ocupados: 01..56 ocupados; este es el 57
padre: plans/56 - jerarquia-seccion-subseccion-plantilla-r1.md
madre_doctrina: .opencode/project-context.md (plantilla autoridad; values-only; guard anti-formula; tolerancia ±0,5; prohibido <v> stale como oraculo) + plans/49 (A1-nucleo + A1-interior, origen del ancla) + plans/54 (CA0 derivacion obligatoria; R-ANCLA-EXTEMP) + plans/56 (§7 R-ANCLA-EXTEMP/O1/O2/O3, §8 corridas y DetRetri) + plans/40 (W-1 evalua, no lee caches) + plans/31 (gate workbook-vs-dominio)
caso_rojo: corrida CLI real 2026101 --cinco-ase — completa lectura, coherencia multi-ASE, validacion DetRetri vs R10, escritura del workbook e INTERVENTORIA de los 5 bloques; muere en el ultimo paso (solo lectura): `[16:29:40] Leyendo oraculo de validaciones cruzadas de la salida (read-only)...` → `ERROR [ERR-PLANTILLA]: La celda-ancla 'Reporte Componentes R1!F682' (EXTEMP del ASE 5) trae formula y el contrato T0 la congelo como literal. (Plan 49 A1-interior: tercera forma = ERR-PLANTILLA; una v3 exige re-congelar el mapa con su T0.)`
gate_entrada:
  plantilla_v2: NO SE TOCA (SHA 0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB — verificar intacto antes/despues)
  plantilla_v1: NO SE TOCA (historica; SHA 3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214 intacto antes/despues)
  alcance_aprobado: H0 read-only por bytes primero (inventario de anclas congeladas en los 5 bloques); cero codigo de produccion en H0; prohibido ampliar sin nuevo OK del Ingeniero
  suite_base: p55-H3-suite.trx (1484/1417/67/0) y p56-H3-suite.trx (1514/1452/62/0) — mismo conjunto de rojos; rojo fuera de lista = STOP
  sin_commits: NUNCA git commit / git push; NUNCA git checkout / git reset
---

# Plan 57 — Anclas de validacion derivadas de la plantilla

> SDD completo PROPOSE → DESIGN → SPEC → TASKS. Fuente unica de verdad para la implementacion una vez aprobado ESTE documento. No se implementa sin aprobacion del Ingeniero (punto de bloqueo: sin su OK no se abre H0; estado inicial PROPUESTO).
> Lectura obligatoria previa: `plans/56` (§7 R-ANCLA-EXTEMP/O1/O2/O3, §8 corridas y DetRetri) + `plans/54` (CA0, lectura unica, residuos) + `plans/49` (A1-nucleo/A1-interior: el origen del ancla; desviacion W-1 del operando) + `plans/40` (W-1 evalua+sella) + `plans/31` (gate workbook-vs-dominio) + `.opencode/project-context.md` (doctrina madre).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Mandato, bloqueo y causa raiz (vinculante)

### 0.1 El bloqueo (verbatim de la corrida real, 2026-10-23 16:29)

La corrida `2026101 --cinco-ase` ya completa lectura, coherencia multi-ASE, validacion `DetRetri` vs R10, **escritura del workbook** e INTERVENTORIA de los 5 bloques. Muere en el **ultimo** paso, que es **solo lectura**:

```text
[16:29:40] Leyendo oraculo de validaciones cruzadas de la salida (read-only)...
[16:29:40] ERROR [ERR-PLANTILLA]: La celda-ancla 'Reporte Componentes R1!F682' (EXTEMP del ASE 5)
trae formula y el contrato T0 la congelo como literal. (Plan 49 A1-interior: tercera forma =
ERR-PLANTILLA; una v3 exige re-congelar el mapa con su T0.)
```

Es el **unico** bloqueo para exit 0. Los `DetRetri` son correctos (14.842.425.136 / 11.302.637.605 / 8.823.787.300 / 10.483.172.762 / 7.550.668.665 = R10) y la plantilla esta intacta.

### 0.2 Causa raiz (medida)

`AnclasExtempR1` (Plan 49, medicion 2026-10-19, `Remuneracion.Core/Models/AnclasExtempR1.cs:45-55`) **congelo como constante** la naturaleza de ciertas celdas-ancla: declaro `Reporte Componentes R1!F682` (y familia) como **literal**. La plantilla v2 vigente (SHA `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB`) trae `F682 = <f>F660+F629-L629</f>`. `ValidacionOracleReader.ResolverAnclasExtempR1` (`Remuneracion.Infrastructure/Excel/ValidacionOracleReader.cs:235-289`, guarda en `:264-269`) exige literal y aborta al encontrar formula.

**Esto es nuestro, no de la plantilla ni de los datos.** Es otro caso de "foto congelada" que se desactualiza: exactamente la clase de problema que el Ingeniero viene prohibiendo ("nada quemado"). Ya consta como residuo `R-ANCLA-EXTEMP` (Plan 54 §7, Plan 55 §7, Plan 56 §7/O3): este plan lo **retira** resolviendolo, no re-congelandolo.

### 0.3 Mandato del Ingeniero (vinculante)

- **Calidad senior**: la solucion debe **eliminar la tabla congelada**, no re-congelar otra foto a mano.
- **Cero cambios a la plantilla**: ni formulas, ni filas, ni geometria. SHA v2/v1 antes/despues.
- **Cero quemado**: nada de listas por celda, por ASE, por empresa ni por quincena.
- El guard que impide escribir sobre celdas con `<f>` **se mantiene intacto** (protege la plantilla).

---

## 1. PROPOSE

### 1.1 Intent

Que el oraculo de validaciones cruzadas (`ValidacionOracleReader` + `ValidadorBasico`, Plan 49 A1) resuelva sus celdas-ancla **derivandolas de la plantilla en runtime** —posicion por la referencia que la propia plantilla declara, naturaleza (formula o literal) leida del workbook en el momento— de modo que las corridas reales de los 3 periodos terminen con exit 0 **sin mover un peso**, sin tabla congelada, sin direcciones congeladas, sin tocar `<f>`, y con el gate absoluto intacto en severidad.

### 1.2 In Scope (H0 + H1 + H2 + H3, prohibido ampliar sin nuevo OK del Ingeniero)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **H0/T0** | Inventario por bytes read-only (cero codigo de produccion): **todas** las celdas-ancla hoy congeladas + su naturaleza real en la v2 (`<f>`? literal? texto de la formula?) en los 5 bloques; veredicto de homogeneidad del patron | Fixture verde + doc de evidencia; puerta H0→H1 |
| **H1** | Resolucion en runtime: posicion derivada + naturaleza leida del workbook (formula → evaluar con los valores del propio workbook; literal → leer); tabla congelada **retirada**; mensajes accionables nuevos | CA-A1..CA-A6 verdes en corrida real |
| **H2** | Cobertura: tests por bloque × periodo (5 × 3) + candado anti-quemado (estatico + injerto) + prueba de que el guard anti-formula del escritor sigue abortando | Suite verde + set-diff exacto |
| **H3** | Doctrina en `.opencode/project-context.md` (solo lo probado) + cierre de este plan; `R-ANCLA-EXTEMP` retirado | Diff de docs citado; cero produccion |

### 1.3 Non-goals / out (vinculante)

- No tocar la plantilla (SHAs v1/v2 antes/despues; cero `<f>`, cero `InsertRow`/`DeleteRow`).
- No relajar el guard anti-formula del escritor (`EscribirValorNumerico` → `ERR-PLANTILLA` intacto; CA-A6 lo pincha).
- No corregir EXTEMP/R2 ni reclasificar divergencias ajenas (Planes 49/50/53 y sus deudas intactas; `ConsolidadoAse.Extemp` intacto — desviacion W-1 del Plan 49).
- No reabrir la semantica de agregados (Plan 25 D-C/D-F), la lectura unica R1 (Plan 54), el resolvedor por etiquetas (Plan 55) ni la jerarquia A/B (Plan 56).
- **Prohibido tocar** `EscritorValoresSlotR1`, `PreparadorBasePipeline`, `ProcesadorPeriodo.cs:198-217` (`R-RECORTE-Q2`), `WorkbookLeafCellMap*`, `ResolvedorConciliacionEmpresa`, `ValidadorBasico`/`WorkbookLeafCoherence` — salvo que el H0 demuestre con archivo:linea que el oraculo vive ahi; en ese caso justificarlo y limitar el cambio a la resolucion del ancla (ver §2-D5).
- Commits, UI, CLI, docs de usuario. NUNCA `git commit` / `git push`. **NUNCA `git checkout` / `git reset`** (incidente plans/55 §8).

### 1.4 Riesgos con mitigacion

| ID | Riesgo | Mitigacion |
|---|---|---|
| R-RECONGELAR | "Arreglar" cambiando el mapa a `F682 = formula` (otra foto congelada) | §2-D1/D2: el mapa se **retira**, no se re-congela; candado §2-D6 + CA-A4; H0 lista **todas** las anclas para que ninguna quede congelada en otro sitio |
| R-TAUTOLOGIA | Evaluar la formula convierte el gate en "el workbook dice lo que el workbook dice" | §2-D4: el veredicto (nucleo) sigue siendo **visible-del-workbook vs esperado-del-dominio**; la evaluacion solo resuelve que muestra el workbook; el interior queda como evidencia diagnostica, no como veredicto |
| R-ORACULO-V | Usar `<v>` stale de celdas con `<f>` como oraculo (leccion Planes 28/29/40) | Metodo H0 y resolucion: texto-`<f>` + literales + fuente leida; `<v>` de celda con `<f>` nunca es oraculo (S0 G-WORKBOOK) |
| R-PLANTILLA | El fix toca una `<f>` o la geometria | SHAs + G-FORMULA + guard `ERR-PLANTILLA` intacto (CA-A6); H0 read-only |
| R-SILENCIO | Ancla no resoluble se vuelve 0 silencioso o se inventa celda | §2-D5: error accionable + alto, **cero escritura** (el reader ya es read-only; el fallo ocurre antes de cualquier veredicto) |
| R-DETRI | El cambio mueve el `DetRetri` | CA-A3: `DetRetri` = R10 al peso en los 15 cruces; mismatch ⇒ STOP |
| R-QUEMADO | Reintroducir direcciones/ids por la puerta de atras | Candado de dos capas §2-D6 + CA-A4 + grep de cierre |

---

## 2. DESIGN

### D1. Metodo H0/T0 por bytes (read-only bloqueante, cero codigo de produccion)

**Forma normativa:** BCL puro (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`), sin Excel/COM, sin OpenXML SDK, sin ExcelDataReader. Rutas explicitas a la v2 + a los tres periodos (2026071, 2026072, 2026101); prohibido `Insumos.Raiz()` en lo nuevo. Cero escritura sobre fuentes y plantillas. Cero codigo de produccion en H0: solo fixture T0 + documento de evidencia `Docs/T0-Plan57-Anclas-Derivadas.md`.

Por cada bloque (5) y cada ancla hoy congelada:

1. **Inventario de lo congelado:** toda entrada de `AnclasExtempR1.Mapa` (`AnclasExtempR1.cs:45-55`: `F682`, `F702` + las que el grep H0 encuentre —buscar tambien `F702` y cualquier otra del contrato Plan 49, mas `CeldaConsolidadoExtemp`/`PrimeraFilaConsolidadoExtemp` y las 10 `<f>` TOT_OPT/EXTEMP declaradas en `R1Q2ProtectedPorAse` por el Plan 54 H1.7). **Debe quedar claro cuantas son** (cota del encargo: familia `F682`/`F702`; el H0 la cierra exacta, no supuesta).
2. **Naturaleza real en la v2:** para cada una —¿trae `<f>`? ¿literal? ¿texto exacto de la formula?— mas existencia de sus celdas de seccion (`F660/F629/L629/F659/F628/L628` y las que correspondan a cada bloque).
3. **Los 5 bloques, no solo el ASE 5:** el mismo inventario en ASE 1..4 (el comentario del mapa afirma que ASE1-4 "traen `<f>` y el resolutor las evalua solo" — el H0 lo convierte en prueba por bytes o lo corrige).
4. **Veredicto de homogeneidad:** ¿el patron es homogeneo (las 5 anclas EXTEMP son `<f>` de la forma `Σ Aplic − Esp(primera)`, cf. Plan 54 F6) o hay bloques con literal real? El veredicto decide la forma de D2; si aparece forma no prevista ⇒ STOP + escalado (no se inventa rama).

**Puerta H0→H1:** inventario cerrado (numero exacto de anclas + naturaleza por bloque) + fixture verde. Sin H0 verde no se abre H1.

| Candidata | Veredicto | Motivo |
|---|---|---|
| **M-A: T0 bytes de las anclas en 5 bloques + retiro del mapa + resolucion derivada** | **Elegida** | Unica que cumple el mandato (eliminar, no re-congelar); precedente: T0-Plan50/52/54/56 |
| M-B: re-congelar el mapa con `F682 = formula` | Descartada | Es otra foto congelada; viola el mandato §0.3 y R-RECONGELAR; el proximo cambio de plantilla la vuelve a romper |
| M-C: eximir el interior (solo nucleo) para pasar | Descartada | Apaga un diagnostico que el Plan 49 ratifico (doble gate §7.1); convierte evidencia en silencio; viola E1-1/E1-2 (Plan 40) |

### D2. Resolucion en runtime (decision)

**Forma normativa del cambio** (objeto: `ValidacionOracleReader.ResolverAnclasExtempR1` + `AnclasExtempR1`; consumo y juicio en `ValidadorBasico` **intactos** en severidad y formato de veredicto):

1. La naturaleza del ancla se **lee del workbook en el momento**: ¿la celda trae `<f>` efectiva (`TextoFormulaEfectiva`, `ValidacionOracleReader.cs:264`) → **evaluar la formula con los valores del propio workbook** (vía `ResolutorValorValidacion.TryEvaluar`, el mismo evaluador del Plan 40 — sin codigo nuevo de evaluacion); ¿literal? → leerlo. **Sin tabla, sin lista, sin ids.** La guarda `:264-269` que aborta ante `<f>` **se retira** (era la foto congelada); el resto de guardas (ancla ausente con secciones presentes, irresoluble) **permanecen**.
2. La expresion de secciones deja de venir del mapa: se deriva de la **formula efectiva leida** (el texto-`<f>` de la propia celda ancla). No hay `ExpresionSecciones` congelada; la evidencia (`ValorAncla` vs `ValorSecciones`) se computa evaluando la formula leida y sus terminos contra lo escrito.
3. `AnclasExtempR1` como tabla de celdas/expresiones **se retira** del camino productivo (borrado con cita, o marca `LEGACY` + retiro si otro consumidor vivo lo exige — el H0/T0.5 lo inventaria; por defecto: borrado). Conservable unicamente lo que no es foto: los nombres de hoja (`Hoja`, `HojaConsolidado`) si el H1 los necesita como constantes de hoja —y aun asi, preferir derivarlos del workbook donde el lector ya los resuelve—.
4. Geometria pura (D-H): el camino del oraculo conoce **etiquetas y referencias leidas**, nunca identificadores de bloque, nombres de empresa, quincenas ni direcciones congeladas (ver candado §2-D6).

### D3. Localizacion del ancla por etiqueta, no por direccion (decision)

Si hoy las anclas estan indexadas por celda (`PorAse(aseId)` → lista de `(Celda, Expresion)`), deben resolverse por el mismo criterio de la plantilla que el resto del proyecto:

1. **Via preferente — seguir la referencia que la plantilla declara:** el ancla EXTEMP del bloque es la celda que referencia el `<f>` de `CONSOLIDADO_TOTAL RECAUDO!D{47+k}` (leido del workbook: p. ej. `D51 → 'Reporte Componentes R1'!F682`). La posicion se **lee**, no se declara: el reader parsea el texto-`<f>` del consolidado (una sola referencia a `Reporte Componentes R1`) y esa es el ancla. Vale para los 5 bloques por la misma ruta (D47→ASE1 … D51→ASE5), sin tabla por ASE.
2. **Ancla interior secundaria** (la familia `F702`, consumida por `REMUNERACION_*`): misma regla — seguir la referencia que la declara (el `<f>` consumidor leido del workbook) o, si el H0 demuestra que no hay consumidor `<f>` estable, localizarla por **etiqueta del bloque** (`EXTEMPORANEO` en la columna D del bloque, o la jerarquia A/B del Plan 56: identidad `(A, B, ocurrencia)` + cruce `(C, D, E)`). El H0 fija cual de las dos vias con bytes; **prohibido reintroducir direcciones congeladas** en cualquiera de ellas.
3. **Prohibiciones:** listas por celda/ASE, literales de celda en el camino del oraculo, ramas por quincena, fallback posicional ("la celda mas cercana"), inventar celda destino.

### D4. Que valida (normativo — anti-tautologia)

El oraculo debe seguir comparando el **valor visible de la salida** contra el **esperado del dominio** (derivado de la fuente). Evaluar la formula no puede convertirse en un chequeo tautologico:

1. **Veredicto (nucleo, Plan 49 A1-nucleo):** `EXTEMP resuelto desde lo escrito` (CONSOLIDADO `D{47+k}`, siguiendo su `<f>` hasta el ancla y evaluando con los valores del workbook) **vs** `WorkbookLeafInputsR1.ExtemporaneoEsperadoPorAse` (agregado del dominio por firma — desviacion W-1 del Plan 49, intacta). Operando izquierdo = workbook; operando derecho = dominio derivado de la **fuente** por el camino independiente del reader (`MapearR1`, Plan 54). La evaluacion de la formula solo resuelve **que muestra** el workbook; no aporta el esperado. **Independencia intacta por construccion:** una perdida comun a lectura-y-escritura seguiria visible porque el esperado no sale del workbook.
2. **Evidencia diagnostica (interior, Plan 49 A1-interior):** `valor del ancla vs suma de sus secciones` (ambos del workbook) es **consistencia interna**, no veredicto: localiza donde se perdio el importe (hoja + celda ancla) cuando el nucleo dispara, pero **nunca convierte un nucleo rojo en verde ni un verde en rojo**. Un interior que no cierra con nucleo verde es hallazgo documentado (falla ruidosa con el mensaje §2-D5), no falso verde.
3. **Lo que NO es oraculo:** cachés `<v>` de celdas con `<f>` (Planes 28/29/40); la intencion del dominio como sustituto de lo escrito (Plan 29: el sellado es read-back, y aqui no hay sellado — el oraculo es read-only).

### D5. Mensajes y comportamiento ante lo irresoluble (normativo)

- Si un ancla no se puede resolver (sin referencia declarada, referencia a hoja/celda inexistente) o la formula no evalua (termino irresoluble desde lo escrito): error accionable que nombra **concepto + bloque + hoja** —p. ej. `ASE {n} · Reporte Componentes R1 · EXTEMPORANEO del bloque: ...`— con esperado vs obtenido donde aplique, codigo `ERR-PLANTILLA`, **y cero escritura** (el reader es read-only; el fallo ocurre antes de emitir veredicto alguno; nada del workbook se modifica).
- Prohibido el texto viejo ("el contrato T0 la congelo como literal"): el mensaje describe el estado del workbook, no la historia del codigo. Jerga de codigo prohibida en el texto (doctrina Plan 26 D-E); el codigo de error viaja en el prefijo.
- **Cambio autorizado (no es "tocar", es el objeto):** `ValidacionOracleReader.ResolverAnclasExtempR1` (+ retiro de `AnclasExtempR1` como tabla) bajo §2-D2/D3. Si el H0 demuestra con archivo:linea que el oraculo vive tambien en `ValidadorBasico`/`WorkbookLeafCoherence` (p. ej. el juicio interior `:627-669`), justificarlo y limitar el cambio a consumir la evidencia derivada (severidad y formato de veredicto intactos). Todo lo demas de §1.3 intacto (grep de cierre).

### D6. Candado anti-quemado (dos capas, nuevo)

Patron Plan 54/55/56, adaptado a este plan:

1. **Estatico:** el camino del oraculo (`ValidacionOracleReader` + lo que quede de `AnclasExtempR1` + el juicio interior de `ValidadorBasico`) no contiene listas de anclas por celda/ASE, literales de celda-ancla, nombres de empresa, ni ramas por quincena. Allow-list cerrada (seguir-referencia + `TextoFormulaEfectiva`/`TryEvaluar` + nombres de hoja) + **prueba de injerto que debe enrojecer** (una entrada de ancla por celda/ASE o una direccion congelada injertadas a proposito en el camino enrojecen el test).
2. **Conductual:** el mismo workbook resuelto para los 5 bloques produce evidencia por la **misma ruta** (misma referencia seguida, misma evaluacion); mas prueba de injerto (patron Plan 54/55/56): reordenar o renombrar lo irrelevante no cambia el veredicto, y una formula injertada con termino falso dispara el nucleo con el Δ exacto.

Sin candado verde no hay H3.

### Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Reader = evidencia read-only derivada (resuelve que muestra el workbook); Validador = juicio (intacto); mapa retirado (el dato congelado desaparece); Writer = values-only intacto |
| **OCP** | ✅ Ancla futura (otra celda, otra expresion, otra naturaleza) absorbida por la referencia leida, nunca por fila nueva ni por rama; celda nueva = etiqueta/referencia del template, nunca `if (ase/empresa/quincena)` |
| **LSP/DIP** | ✅ Contratos `IValidacionOracleReader`/`IValidador` intactos; campos del snapshot aditivos o intactos; cambios detras de las mismas abstracciones |
| **Best practices** | ✅ Contract-first (H0 congela inventario + naturalezas antes del juicio); fail-fast accionable (concepto+bloque+hoja, cero escritura); tolerancia ±0,5 solo en comparacion; values-only; D-H (cero celda/ASE/empresa/quincena en resolucion); CRLF; sin commits; prohibido `git checkout`/`git reset`; prohibido `<v>` de celda con `<f>` como oraculo |
| **Performance** | ✅ H0 acotado BCL ZIP+XML + fixtures existentes; resolucion O(anclas × terminos) por bloque en memoria sobre el workbook ya abierto; regens a temp fresca; cero workbooks extra en runtime |

---

## 3. SPEC (criterios numerados + gates no negociables)

### S0 — Gates transversales (toda H)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa trx + **set-diff vs `p55-H3-suite.trx` (1484/1417/67/0) y `p56-H3-suite.trx` (1514/1452/62/0)**: altas/bajas una por una con causa citada; **rojo nuevo fuera de la lista H0 ⇒ STOP**. 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados, 0 tolerancias ampliadas; re-anclajes solo con evidencia.
- G-SHA: v1 `3086AC31…C5214` + v2 `0F192AD2…7F2DAB` intactos antes/despues de cada H.
- G-FORMULA: `<f>` byte-identica en todo el diff (cero `new CellFormula` fuera de `EscribirFormulaPreservando`, cero `InsertRow`/`DeleteRow` fuera de Phase 0). El guard anti-formula del escritor intacto (CA-A6).
- G-D-H: cero listas de anclas por celda/ASE, cero direcciones congeladas, cero nombres de empresa, cero ramas por quincena en el camino del oraculo (grep de cierre + candado §2-D6).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale) o la fuente ya leida; **cero workbooks extra en runtime**.
- G-TRAZA: cada error nuevo nombra concepto + bloque + hoja (+ esperado vs obtenido donde aplique).
- G-PROC: cero commits; **cero `git checkout` / `git reset`**; CRLF sin mezclar en lo tocado.
- G-H0-READONLY: H0 no contiene codigo de produccion (solo fixture H0 + doc de evidencia); el diff de `Remuneracion.Core/` + `Remuneracion.Infrastructure/` en H0 es vacio salvo tests `*T0Plan57*` / `*AnclaDerivada*`.

### CA — Criterios de aceptacion (vinculantes; con derivacion obligatoria, CA0 del Plan 54)

- CA0: **todo numero de criterio viaja con su derivacion** (`fuente × <f>` v2 + hoja/celda). Un numero sin derivacion no se acepta como criterio.
- **CA-A1**: la corrida CLI `2026101 --cinco-ase` **termina con exit 0** (sin `ERR-PLANTILLA`, sin `ERR-SIN-SLOT`, sin `ERR-VALIDACION`), con workbook validado. Derivacion: oraculo resuelto por referencia leida (§2-D2/D3) + `DetRetri` = R10 (CA-A3).
- **CA-A2**: idem para `2026071` y `2026072` (misma ruta, sin ramas por quincena).
- **CA-A3**: `DetRetri` = R10 en los 5 bloques de los 3 periodos, **sin mover un peso** (los 15 valores, con derivacion `fuente × <f>` v2):

| Periodo | ASE1 | ASE2 | ASE3 | ASE4 | ASE5 |
|---|---|---|---|---|---|
| 2026071 | 16.758.167.585 | 12.227.292.053 | 10.127.195.599 | 10.553.776.376 | 8.543.663.209 |
| 2026072 | 17.450.228.673 | 20.516.143.970 | 15.221.896.467 | 7.179.595.396 | 12.073.344.662 |
| 2026101 | 14.842.425.136 | 11.302.637.605 | 8.823.787.300 | 10.483.172.762 | 7.550.668.665 |

  (Fuente: Plan 56 §8; el H0/T0.4 los re-deriva `fuente × <f>`; mismatch ⇒ STOP.)
- **CA-A4**: cero listas de anclas por celda/ASE y cero direcciones congeladas en el camino del oraculo (candado verde + injerto que enrojece, §2-D6).
- **CA-A5**: plantilla byte-identica salvo los 11 tokens de periodo; SHAs v1/v2 intactos.
- **CA-A6**: el guard anti-formula del escritor sigue intacto (test que prueba que escribir sobre `<f>` sigue abortando con `ERR-PLANTILLA`).

### S1 — Criterios H0/T0 (read-only, cero codigo de produccion)

- S1.1: inventario cerrado de **todas** las celdas-ancla hoy congeladas (incluidas `F682`, `F702` y cualquier otra del contrato Plan 49) con su naturaleza real en la v2 (`<f>` / literal / texto de la formula) en los 5 bloques; numero exacto + veredicto de homogeneidad.
- S1.2: referencia declarante por bloque (el `<f>` de `CONSOLIDADO!D47:D51` → ancla; consumidor del ancla interior) leida por bytes; via de localizacion fijada (§2-D3.1 vs §2-D3.2).
- S1.3: cierre aritmetico por bloque contra el R10 de cada periodo, al peso, con derivacion `fuente × <f>` (re-deriva CA-A3).
- S1.4: lista cerrada de Δ de suite esperados (por defecto: ninguno — el H0 es read-only; cualquier rojo nuevo en H0 ⇒ STOP). **Sin H0 verde no se abre H1.**

### S2 — Criterios H1 (implementacion)

- S2.1: resolucion derivada implementada (§2-D2/D3); tabla congelada retirada del camino productivo; D-H intacto (grep + CA-A4).
- S2.2: mensajes nuevos (concepto + bloque + hoja, cero escritura) implementados; texto viejo ("el contrato T0 la congelo") eliminado del camino (grep).
- S2.3: CA-A1 + CA-A2 + CA-A3 reproducidos en corridas reales de los 3 periodos (SHAs + G-FORMULA + 0 warnings).
- S2.4: archivos prohibidos intactos (§1.3/§2-D5, grep de cierre salvo justificacion H0 con archivo:linea); juicio del nucleo intacto en severidad (un operando es el dominio).

### S3 — Criterios H2 (regresion + cobertura)

- S3.1: tests por bloque × periodo (15: 5 bloques × 2026071/2026072/2026101) que ejercitan ancla formula y ancla literal (la v2 trae formula en los 5 —el caso literal se pincha con fixture sintetico + la via literal del codigo—); test de mensaje irresoluble (concepto+bloque+hoja, cero escritura); candado §2-D6 verde en sus dos capas (incluida prueba de injerto en rojo).
- S3.2: set-diff vacio de rojos vs las dos lineas base salvo altas citadas; CA-A6 verde (guard anti-formula del escritor aborta).
- S3.3: code-review independiente **APROBADO** (0 bloqueantes): spec compliance + calidad + greps D-H + SHAs.

### S4 — Criterios H3 (docs/cierre)

- S4.1: doctrina en `.opencode/project-context.md` (anclas derivadas; retiro de `R-ANCLA-EXTEMP`; candado); nota de retiro del residuo sin reescribir historia.
- S4.2: este plan como fuente de verdad; desviaciones con fecha/motivo/evidencia; cero codigo de produccion en H3.

---

## 4. TASKS (fases H0..H3 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero codigo de produccion)

Fixture: `Remuneracion.IntegrationTests/AnclasDerivadasT0Plan57Tests.cs` (BCL ZIP+XML, rutas explicitas a v2 + 2026071/2026072/2026101; prohibido `Insumos.Raiz()` en lo nuevo). Evidencia: `Docs/T0-Plan57-Anclas-Derivadas.md`.

| ID | Tarea | Verifica |
|---|---|---|
| T0.1 | Inventariar **todas** las anclas congeladas: entradas de `AnclasExtempR1.Mapa` + `CeldaConsolidadoExtemp`/`PrimeraFilaConsolidadoExtemp` + 10 `<f>` de `R1Q2ProtectedPorAse` + grep de cualquier otra celda-ancla del contrato Plan 49 (buscar `F702` y familia); numero exacto | S1.1 |
| T0.2 | Naturaleza real en la v2 por bloque (5 bloques): ¿`<f>`? ¿literal? ¿texto exacto de la formula? + existencia de las celdas de seccion; veredicto de homogeneidad del patron | S1.1 |
| T0.3 | Referencia declarante por bytes: texto-`<f>` de `CONSOLIDADO!D47:D51` → ancla por bloque + consumidor del ancla interior; fijar via §2-D3.1 vs §2-D3.2 | S1.2 |
| T0.4 | Cierre aritmetico por bloque vs R10 de cada periodo al peso, con derivacion `fuente × <f>` (re-deriva los 15 `DetRetri` de CA-A3) | S1.3 |
| T0.5 | Inventario cerrado de consumidores de `AnclasExtempR1` (reader, `ValidadorBasico.cs:627-669`, tests) con veredicto por consumidor (retirar / aislar-legacy / re-anclar-test) | S2.4 (anticipado) |
| T0.6 | Lista cerrada de Δ de suite + SHAs v1/v2 antes/despues de H0 + `<f>` de cierre intactas | S1.4 |

- **Puerta H0→H1:** fixture verde + S1.1-S1.4 cerrados + APROBADO del Ingeniero. Si el patron no es derivable o algun cierre no cuadra ⇒ STOP + escalado (no se inventa celda, no se re-congela, no se abre H1).

### H1/F1 — Resolucion derivada + retiro de la tabla (traza a §2–§3)

| ID | Tarea | Traza |
|---|---|---|
| H1.1 | Resolver el ancla por referencia leida + naturaleza en runtime (formula → evaluar con lo escrito; literal → leer); retirar la guarda "trae formula = ERR-PLANTILLA"; derivar la expresion de secciones del `<f>` leido | D2, S2.1 |
| H1.2 | Localizar por etiqueta/referencia, no por direccion (§2-D3); retirar `AnclasExtempR1` como tabla del camino productivo (borrado citado o LEGACY justificado por T0.5) | D3, S2.1 |
| H1.3 | Mensajes nuevos (concepto + bloque + hoja, cero escritura); eliminar el texto viejo del camino | D5, S2.2 |
| H1.4 | e2e 3 periodos (5 bloques): CA-A1 + CA-A2 + CA-A3 (exit 0, `DetRetri` = R10 al peso) | S2.3 |
| H1.5 | Greps de cierre §1.3/§2-D5 + SHAs + G-FORMULA + 0 warnings | S2.4 |

- **Puerta H1→H2:** S2 verde + CA-A1..CA-A6 encaminados + 0w + SHAs + G-FORMULA + greps D-H.

### H2/F2 — Regresion, cobertura y evidencia

| ID | Tarea | Verifica |
|---|---|---|
| H2.1 | 15 tests bloque × periodo (formula + literal) + test de irresoluble (mensaje + cero escritura) + candado §2-D6 (estatico allow-list + injerto en rojo; conductual misma-ruta) | CA-A4, S3.1 |
| H2.2 | Test CA-A6: escribir sobre `<f>` sigue abortando (`ERR-PLANTILLA` del escritor intacto) | CA-A6, S3.2 |
| H2.3 | Suite trx + set-diff vs `p55-H3-suite.trx` y `p56-H3-suite.trx`; fuera de lista = STOP | S0 G-SUITE |
| H2.4 | Code-review independiente (`code-reviewer`, este plan como spec): **APROBADO** requerido | S3.3 |

- **Puerta H2→H3:** set-diff exacto + CA-A1..CA-A6 + review aprobado.

### H3/F3 — Docs/cierre

| ID | Tarea | Verifica |
|---|---|---|
| H3.1 | Doctrina en `.opencode/project-context.md` (solo lo probado: anclas derivadas, retiro de `R-ANCLA-EXTEMP`, candado) | S4.1 |
| H3.2 | Cierre de este plan (§8) con desviaciones fechadas si las hubo; cero codigo de produccion | S4.2 |

**Orden sugerido:** H0 → H1 → H2 → H3. Regresion continua (suite + SHAs en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero); **prohibido `git checkout`/`git reset` sin copia previa** (incidente Plan 55 §8).

---

## 5. Riesgos (consolidado + gates por fase)

Ver §1.4 (R-RECONGELAR, R-TAUTOLOGIA, R-ORACULO-V, R-PLANTILLA, R-SILENCIO, R-DETRI, R-QUEMADO).

- **Gate H0→H1:** H0 verde + S1.3/S1.4 sin STOP + APROBADO del Ingeniero.
- **Gate H1→H2:** S2 + CA-A1..CA-A3 + 0w + SHAs + G-FORMULA + greps.
- **Gate H2→H3:** set-diff exacto + CA-A4..CA-A6 + review APROBADO.
- **Gate H3:** diff de docs citado; cero produccion.

---

## 6. Clarification Gate (requiere aprobacion del Ingeniero; sin esto no se abre H0)

1. **H0 primero:** ¿se ratifica H0 read-only bloqueante (cero codigo de produccion) que inventarie **todas** las anclas congeladas + su naturaleza real en la v2 en los 5 bloques, con veredicto de homogeneidad, y la regla "patron no derivable o cierre que no cuadra ⇒ STOP"?
2. **Retiro vs re-congelamiento:** ¿se ratifica **eliminar** la tabla congelada (`AnclasExtempR1` fuera del camino productivo) frente a re-congelarla con la formula actual?
3. **Resolucion y localizacion:** ¿naturaleza leida del workbook en el momento (formula → evaluar con lo escrito; literal → leer) + posicion por referencia/etiqueta leida (§2-D3), con prohibicion de reintroducir direcciones congeladas?
4. **No-tautologia:** ¿el veredicto sigue siendo visible-del-workbook vs esperado-del-dominio (interior solo diagnostico), segun §2-D4?
5. **Candado y linea base:** ¿candado anti-quemado de dos capas (estatico allow-list + injerto en rojo; conductual) + set-diff contra `p55-H3-suite.trx` (1484/1417/67/0) y `p56-H3-suite.trx` (1514/1452/62/0), con rojo fuera de lista = STOP?
6. **Puerta:** sin 1-5 no se abre H0; sin H0 verde no hay H1; sin candado ni review no hay H3.

Sin estas aprobaciones el plan no avanza de PROPUESTO a APROBADO, y sin APROBADO no se abre H0.

> **RATIFICADO 2026-10-23 por el Ingeniero (calidad senior) — gate §6 completo (1-6):** H0 read-only bloqueante con inventario de todas las anclas y veredicto de homogeneidad · **eliminar** la tabla congelada (no re-congelarla) · naturaleza leída en runtime + posición por referencia/etiqueta leída · no-tautología (visible-vs-dominio; interior solo diagnóstico) · candado de dos capas + líneas base `p55-H3` y `p56-H3` · puertas H0→H1→H2→H3.
> **Condición del Arquitecto (vinculante):** cero tablas de anclas por celda/ASE, cero direcciones congeladas y cero ramas por quincena en el camino del oráculo; la plantilla y su guard anti-fórmula no se tocan. Si H0 revela heterogeneidad no prevista ⇒ **STOP**, no se inventa rama.

---

## 7. Residuos declarados (fuera de alcance — NO implementar)

**R-RECORTE-Q2 (vivo, sin tocar):** `Remuneracion.Core/Services/ProcesadorPeriodo.cs:198-217` mantiene la omision de la conciliacion por empresa en Q2. Deuda `R-MAPA-V2-R2-ASE4` (Plan 50 §8). Gate de reactivacion: Plan 55 §D6. No se toca aca (§1.3).

**R-SIGMA-Q2 (vivo, no se re-litiga):** el gate Sigma no es descomposicion valida del visible en Q2 (T0-Plan50 §4.2). No se toca aca.

**R-ETIQUETAS-SIN-DESTINO (marco aplicable, Plan 55 §7):** si el cruce de etiquetas encuentra un exceso real futuro, aplica error accionable + alto sin escribir; prohibido inventar celdas destino.

**O1 — Candado estatico parcial (deuda del review del Plan 56, con gate, intacta):** el candado §2-D5 del Plan 56 cubre en `EscritorValoresSlotR1` solo el bloque de planificacion, no el camino historico `TryPlanificar`/`MensajeSinSlotJerarquia`. No se toca aca.

**O2 — Registro durable del cero documentado (deuda del review del Plan 56, con gate, intacta):** el «registro durable» es hoy `Log.Warning` + `CerosDetalle` en memoria; no persiste como artefacto de corrida. No se toca aca.

**R-ANCLA-EXTEMP — LO RESUELVE ESTE PLAN (no es residuo al cierre):** al cerrar H3, `AnclasExtempR1` como tabla congelada queda retirada y el oraculo deriva sus anclas de la plantilla; la entrada se da de baja con cita a la evidencia H2. Si H0 demuestra forma no derivable ⇒ STOP y el residuo sigue vivo (no se finge cierre).

---

## 8. Estado

**APROBADO 2026-10-23** (gate §6 ratificado). H0 habilitado y bloqueante; H1 condicionada al T0 de homogeneidad. Padre: Plan 56 (CERRADO). Madre doctrina: `.opencode/project-context.md` + Plans 49/54/56 (+ 40/31 para el oraculo).
