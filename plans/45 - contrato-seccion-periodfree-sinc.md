---
titulo: Plan 45 - Contrato de seccion PERIOD-FREE para filas sin-C + cero documentado (hijo del Plan 44)
estado: CERRADO 2026-10-16 — H0 (T0 multi-periodo) + H1 (regla sin-C + cero documentado, Q1-neutral) + H2 (review APROBADO-CON-OBSERVACIONES con O-1 formalizada; set-diff VACIO 39=39 en p45-h2b.trx) + H3 (docs). Follow-ups declarados: O-2/O-3. Puertas abiertas en el padre 44: Via A/B (F2) y auditoria de la propietaria.
fecha: 2026-10-16
numero: 45
planes_ocupados: 01..44 ocupados; este es el 45
padre: plans/44 - llenado-r1-fuente-canal-concepto-componentes.md (Enmienda E-44.1 vigente e intacta)
gate_entrada:
  padre_44_E-44_1: F1-MULTICOLUMNA DONE y verificado por el orquestador 2026-10-16 (build 0w, 22/22, anchor-neutral, set-diff vacio); regla sin-C + cero documentado diferidos a ESTE plan hijo
  decision_ingeniero_2026-10-16: F1 del padre CIERRA en multicolumna; ESTE plan hijo resuelve la regla sin-C + cero documentado
  q1: NO SE TOCA (ni codigo, ni tests verdes como SeccionEspejoR1Tests.T1_Ase2Q1; el T0 solo LEE bytes de Q1 read-only; cualquier cambio Q1 exige autorizacion explicita separada)
  plantilla_oficial: NO SE TOCA (SHA256 3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214 intacto antes/despues)
  catalogo_divergencias: NO SE TOCA (DivergenciasDeclaradasCatalogo intacto; Via A/B queda en el padre/F2)
---

# Plan 45 - Contrato de seccion PERIOD-FREE para filas sin-C + cero documentado (hijo del Plan 44)

> SDD completo PROPOSE -> DESIGN -> SPEC -> TASKS. Fuente unica de verdad para la implementacion una vez aprobado ESTE documento. No se implementa sin aprobacion del Ingeniero.
> Lectura obligatoria previa: plan padre completo + Enmienda E-44.1 + `Docs/T0-Plan44-R1-Fuente-Canal-Concepto.md` (especialmente S2 E8 inversion de fronteras ASE5, S3 colocacion 11/11, S4 hueco Intereses-1.a, S5 F654/F661-escritura, S10 STOP, S12 P1/P2/P3) + `Remuneracion.Core/Models/R1FirmaInterior.cs` (`EsDatoEmpresa` exige C) + `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs` (cursor por grupo, F/L extendido a F..Z en el padre) + `.opencode/project-context.md` (doctrinas e invariantes) + plans/41 (E1-6/E2-5: contrato de seccion Q1-aware diferido aqui).
> Visual Design Intent: N/A (cero UI en este plan).

## 0. Contexto adoptado del padre (no re-diagnosticar; cada gate lo re-verifica por test)

### 0.1 Problema probado en bytes (T0 del padre, adoptado)

- La clave `(firma + tipo/ordinal de frontera)` alinea ASE1-4 pero NO ASE5. Causa citada (T0 S2 E8): la fuente ASE5 trae `Aplic, Mes, Aplic, Mes` (4 fronteras, empieza con Aplicacion) vs `Mes, Aplic, Mes, Aplic, Mes` (5) del oficial y de ASE1-4. No existe emparejamiento puro por `(tipo, ordinal)` que lleve cada fila fuente ASE5 a su slot sin huecos; la forma del bloque manda.
- Casos multi-periodo que el contrato debe absorber sin ramas (geometria citada, no supuesta):
  - Agosto-2026082 ASE5 trae 3 secciones `Mes` (forma distinta de julio-2, que trae 2). La fuente define la forma; el contrato no congela conteos de `Mes`.
  - Q1-ASE2 trae `D=5` en 1.a seccion (fuente Q1 sin filas-frontera `Aplicacion nuevos x reversion`; estructura `Primera / Oportuno#1 / Oportuno#2`, Plan 41 E1-6). El contrato lo CITA en el T0 propio pero NO lo resuelve en codigo (RESTRICCION DURA, ver 0.2).
- Colocacion vigente medida 11/11 (T0 S3): el defecto vive enteramente en el cursor fisico de filas sin C (firma-only). La regla del padre S2.1 NO lo corrige (veto de seccion no aplica a filas sin C) y extenderla ingenuamente rompe ASE5 (la `Vlr Intereses` de 1.a no tiene slot en 1.a oficial; el bloque `Componente/Subsidio` cae en seccion oficial de forma distinta).
- Hueco de plantilla (T0 S4): 1.a seccion oficial sin `Vlr Intereses` en los 5 ASE; fuente ASE5 julio-2 SI trae `O` 402,40 y `T` 7.697,23 con desglose; manual legacy ASE5 SI los registraba (`E506/F506=402,40`, `E509/F509=7697,23`). NO se toca la plantilla; se reporta al dueno (texto T0 S9). Contradiccion bloqueante resuelta por E-44.1: esas filas = cero documentado, no `ERR-SIN-SLOT`.
- Micro-punto F654/F661 (T0 S5): artefacto de ESCRITURA (cursor fisico firma-only), no del lector (una emision por fila: R47 `Subsidio` y R48 `D=H/Total`, ambas `F=-187.855,98`, caen en 654 y 661 por cursor). La regla sin-C SI puede y DEBE taparlo; la condicion de escape del padre S2.2 no aplica.
- Veredictos ya tomados y congelados (no reabrir): T/Y no espejo (writer solo T, Y intacta; resolucion por `(texto, primera ocurrencia)`); `U(fuente)->Z(plantilla)`; `Exonera-*` U/V/W/X intactas; anclas OCCIDENTE/F628/F2 intactas para F2 del padre; Via A/B solo en F2 del padre con numeros frescos post-hijo.

### 0.2 RESTRICCION DURA Q1 (vinculante, sin excepcion)

- Q1 (2026071) NO SE TOCA: ni codigo de produccion que altere su comportamiento, ni tests verdes como `SeccionEspejoR1Tests.T1_Ase2Q1` (ese test documenta el bloqueo Q1 declarado y sigue rojo por su causa hasta su propio contrato).
- El T0 propio de ESTE plan solo LEE bytes de Q1 read-only (BCL ZIP+XML, sin Excel/COM/OpenXML, sin modificar nada) para derivar la regla PERIOD-FREE sobre evidencia real de los 3 periodos (julio-2 + agosto + Q1). Lectura != resolucion: citar Q1 no autoriza cambiar Q1.
- Cualquier cambio de comportamiento Q1 (incluido reverdecer un rojo Q1) exige autorizacion explicita separada del Ingeniero fuera de este plan. El gate de regresion lo verifica por set-diff (cero rojos Q1 nuevos, cero verdes Q1 nuevos sin autorizacion).
- Tampoco se toca la plantilla oficial (SHA256 intacto) ni `DivergenciasDeclaradasCatalogo` (Via A/B queda en el padre/F2 con numeros frescos post-hijo).

### 0.3 Estado heredado del padre (puertas intactas, no reabiertas)

- Enmienda E-44.1 vigente: P2 (contrato de seccion explicito por ASE, citado en bytes; geometria, cero ramas por periodo) + filas sin slot como cero documentado + reporte al dueno. F1 del padre CIERRA en multicolumna (F..Z por encabezado, `U->Z`, T solo-T, decimal exacto, guard anti-formula, cursor por grupo).
- Auditoria manual de la propietaria: PENDIENTE (puerta explicita, no verde por ausencia). El hueco `Vlr Intereses` se reporta al dueno, no se parchea.
- Linea base de suite vigente a la entrada de este plan: **39 Failed / 888 Passed / 927 Total** (mismos 39 fallos; 8 de Q1 declarados; 12 del fixture F0 verdes). Criterio vinculante: mismos 39 fallos salvo veredicto documentado de este plan; set-diff vacio como puerta.

---

## 1. PROPOSE

### 1.1 Intent

Cerrar la regla sin-C pendiente del padre con un contrato de seccion PERIOD-FREE (valido en julio-2 + agosto + Q1 como evidencia, sin ramas por periodo) y Q2-aware por ASE (la alineacion fuente<->oficial vive en el contrato por ASE, no en el tipo/ordinal generico): cada fila fuente sin empresa en C coloca por `(firma + seccion-contrato)` donde la seccion la resuelve el contrato explicito por ASE derivado sobre bytes reales de los 3 periodos; las filas sin slot de la lista documentada quedan en cero documentado + reporte al dueno; `ERR-SIN-SLOT` queda solo para fila con (firma, seccion-contrato) sin slot fuera de la lista, con fail-fast que nombra ASE+cuadro+fila. El micro-punto F654/F661 queda tapado por la regla (artefacto de escritura probado).

### 1.2 In-scope EXPLICITO

| Frente | Alcance | Criterio de salida |
|---|---|---|
| **H0 - T0 multi-periodo propio (read-only, cero codigo de produccion)** | Fronteras y formas por ASE x periodo (julio-2, agosto-2026082, Q1-2026071 read-only) + matriz fuente<->oficial completa por ASE x periodo + regla sin-C derivada sobre bytes reales de los 3 periodos + lista documentada de ceros con valores diferidos citados + aislamiento F654/F661 ya probado re-verificado + formas Q1 citadas sin resolver. | Doc `Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md` + fixture read-only verde que cita bytes de los 3 periodos; regla sin-C y lista de ceros congeladas o STOP declarado. |
| **H1 - Regla sin-C + cero documentado (contract-first)** | Predicado de seccion-contrato por ASE (Core puro, sin I/O, sin periodo) + consumo en `EscritorValoresSlotR1` (grupos por firma+seccion-contrato, cursor por grupo) + cero documentado (lista T0; sin `ERR-SIN-SLOT`, sin simular; Totales ya llevan el dinero) + `ERR-SIN-SLOT` solo fuera de lista con fail-fast enriquecido. Incluye tapado F654/F661. | 11/11 colocaciones medidas ASE5 corregidas por forma + filas con empresa intactas + hueco Intereses-1.a en cero documentado + F654/F661 sin duplicado + build 0 warnings + greps D-H. |
| **H2 - Regresion + paridad (set-diff, goldens, DetRetri-R10)** | Suite completa con trx + set-diff vs linea base + goldens fila-a-fila julio-2/agosto + `<f>` byte-identica en julio-identidad + DetRetri-D vs R10 por ASE + autocontroles W-1 re-verificados como insumo F2 del padre (sin re-decidir Via A/B). | Set-diff sin rojos nuevos (Q1 intacto) + 0 warnings + SHA256 intacto + tabla observado-vs-esperado para F2 del padre. |
| **H3 - Docs** | `project-context.md` (doctrina contrato PERIOD-FREE + regla sin-C + lista de ceros) + nota al manual HU-17 + reporte al dueno (texto listo, sin enviar plantilla). | Diff de docs citado; cero codigo de produccion en H3. |

### 1.3 Non-goals / out (vinculante)

- Q1: frente propio con su propio contrato (Plan 41 E1-6/E2-5 lo difirio AQUI como evidencia, no como fix). Este plan CITA Q1 en el T0 y NO LO RESUELVE en codigo. Ningun test Q1 se toca ni se reverdece.
- Plantilla oficial: NO se edita (ni oficial ni copia fuera de Phase 0 allow-list). Hueco `Vlr Intereses` se reporta, no se parchea.
- `DivergenciasDeclaradasCatalogo`: intacto. Via A/B solo en F2 del padre con numeros frescos post-hijo.
- Otras pestanas (R2-detalle, R4-detalle, SALDOS POR NOTA detalle, AJUSTES-SF-T, BCE, banco, conciliaciones, INTERVENTORIA salvo lo doctrinal, AFaseo-consumidores): follow-up explicito, no se tocan.
- Auditoria manual: pendiente de entrega; paridad contra auditoria queda PENDIENTE (puerta explicita del padre).
- Re-diagnosticar lo probado (inversion ASE5, 11/11, T/Y, U->Z, F654/F661-escritura): adoptado, no re-diagnosticado; cada gate lo re-verifica por test.

### 1.4 Riesgos con mitigacion

| ID | Riesgo | Mitigacion |
|---|---|---|
| R-FORMA | El contrato por-ASE memoriza julio-2 y rompe agosto (3 `Mes` ASE5) o Q1 (`D=5`, sin `Aplicacion`). | Contrato PERIOD-FREE por construccion: T0 multi-periodo deriva la regla sobre bytes reales de los 3 periodos; geometria citada (cero ramas por periodo/empresa); tests multi-periodo (julio-2 + agosto + Q1-read-only) verdes antes de H1. |
| R-Q1-CONTAGIO | Tocar Q1 al implementar la regla (fallback posicional prohibido E1-1 Plan 41). | RESTRICCION DURA 0.2 + gate set-diff Q1 (cero verdes Q1 nuevos sin autorizacion) + grep D-H + `EsDatoEmpresa` intacto. |
| R-ATRIBUCION | Atribuir a colocacion lo que es plantilla (F628 huerfano, F682/F702 literales, F654/F661). | T0 aisla por bytes: grep de consumidores `<f>` (maestros shared incluidos) + traza lector vs escritor por (firma, seccion-contrato, ocurrencia); veredicto citado por celda; sin veredicto no hay fix. |
| R-SILENCIO | El cero documentado tapa un defecto real (dinero sin desglose que si tenia slot). | Lista cerrada por T0 con valores diferidos citados (402,40 / 7.697,23 + desglose + las que liste el T0 por ASE); Totales verificados (O 50.323,46 = 49.921,06 + 402,40; T 314.490 = 306.792,77 + 7.697,23); fuera de lista = `ERR-SIN-SLOT` fail-fast. |
| R-AUDITORIA | Declarar exito sin oraculo (auditoria pendiente). | Puerta explicita: T0 congela geometria + ejemplos; SPEC deja el gate workbook-vs-dominio con la auditoria como oraculo al llegar; oraculos vigentes (fuente por encabezado, goldens, R10, `<f>` byte-identica). |
| R-PLANTILLA | Nueva version del oficial rompe el contrato. | Contrato por etiquetas + secciones citadas, no por direcciones: `ContratoSlotR1T0Tests` + `GeometriaPlantillaEstandarizadaT0Tests` rompen primero (`ERR-PLANTILLA`); re-congelamiento por T0 nuevo, nunca edicion silenciosa. |
| R-FALSO-VERDE | Apagar/sellar/re-etiquetar un gate (criterio E1-1..E1-5 Plan 40). | Prohibido `SeGatea=false`; sub-bloque no evaluable = hallazgo/defecto con frente propio, nunca verde. |
| R-COLUMNA | Escribir columna formula o inexistente (Exonera-*, Y, G..AP con `<f>`). | Guard anti-formula intacto (`CellFormula` -> `ERR-PLANTILLA`); allow-list de columnas por encabezado; Exonera-* intactas; Y intacta (T/Y no espejo). |

---

## 2. DESIGN

### 2.1 Contrato de seccion PERIOD-FREE (clave + resolucion por ASE)

**Clave de fila (localizador puro, Core, sin I/O, sin periodo - D-H):**

```
clave(fila) = ( C-por-sinonimia , D , E , seccion-contrato-por-ASE )
```

- **C-por-sinonimia:** `SinonimosEmpresaR1.SonMismaEmpresa(fuente.C, slot.C)` (Plan 34, catalogo abierto). Cero ramas por nombre de empresa. Normalizacion mayusculas invariantes + sin diacriticos. Rotulo no tabulado solo se iguala a si mismo.
- **D (canal):** igualdad exacta `OrdinalIgnoreCase` tras `Trim` (H, O, T, E, 1, 2, 5; mas `5` Q1-ASE2 citado en T0). Sin sinonimia.
- **E (concepto):** igualdad exacta `OrdinalIgnoreCase` tras `Trim` (`Vlr Servicio`, `Vlr Intereses`, `Total`; mas `Mes` / `Aplicacion nuevos x reversion` / `Subsidio` / `Subs-Cont` / `Componente` / `AFaseo` como roles-frontera). Sin sinonimia.
- **Seccion-contrato-por-ASE:** la seccion de la fila fuente la resuelve el contrato explicito por ASE derivado en H0 (tabla fuente<->oficial citada en el T0 por ASE x periodo), NO el (tipo, ordinal) generico. El contrato es geometria (cero ramas por periodo/empresa): recibe firmas/fronteras y formas, nunca strings de periodo/empresa. `R1FirmaInterior.EsDatoEmpresa` (exige C no vacia) queda INTACTO; `R1FirmaInterior.EsAplicacionPorFrontera` (frontera por DEBAJO, Plan 35-T2) queda INTACTO (otro consumidor, otra semantica). `SeccionEspejoR1.PorFronteraArriba` sigue resolviendo la seccion de filas-dato de empresa y de slots; la EXTENSION a filas sin C vive en el contrato por-ASE (P2 de E-44.1), no en un veto generico.
- **Afinidad de grupo:** el writer agrupa slots por **(firma A/B + clave + seccion-contrato)** (`ConstruirGruposSeccion` extendido); cursor INDEPENDIENTE por grupo (un slot intercalado sigue disponible para una fila emitida despues). `LeerSlots`/`RangoDatosAse` (ASE1 6-72, ASE2 110-210, ASE3 289-355, ASE4 450-550, ASE5 603-676) intactos en geometria; lectura de celdas F..Z por encabezado intacta del padre (primera ocurrencia manda: `Intereses Cte` -> T, Y intacta; `Rec.Intereses` U -> Z).
- **PERIOD-FREE por construccion:** el contrato declara por ASE la alineacion fuente<->oficial que absorbe las 3 formas probadas (julio-2 5/5/5/5/4 con inversion ASE5; agosto ASE5 con 3 `Mes`; Q1 sin `Aplicacion` y con `D=5` en 1.a) sin literales ni ramas por periodo. Un periodo nuevo con forma nueva = nuevo T0 (re-congelamiento citado), nunca rama ad hoc.
- **Q2-aware por ASE:** dentro del contrato, la tabla por ASE distingue la alineacion vigente Q2 (julio-2/agosto como evidencia) de la forma Q1 citada-sin-resolver. El codigo no recibe periodo; el contrato por-ASE es geometria congelada, no logica por periodo.

- **Neutralidad Q1 (O-1 del review H2.1 — formalizada 2026-10-16, decision del Ingeniero opcion 1):** el writer es **plan-then-write a nivel de BLOQUE**: (1) planifica el bloque completo con `porContrato:true`; (2) si el bloque NO es colocable por forma (`TryPlanificar` falla: fila fuera de lista y sin slot), **replanifica el bloque entero** con `porContrato:false` = comportamiento historico Plan 41 (`ClasificarSeccionesBaseline`, cursor firma-only). Nunca escribe a medias (la escritura ocurre solo tras elegir el plan) y cada intento crea grupos/cursors frescos (el segundo no hereda estado del primero). **Racional:** Q1 se preserva por CONSTRUCCION (su forma no es colocable por contrato → cae al pool historico byte-a-byte), no por rama de periodo/empresa (D-H intacto); el cero documentado solo aplica en el path del contrato. **Tradeoff declarado:** un bloque Q2 futuro que tampoco sea colocable por contrato se resolveria por el pool historico del Plan 41 (no falla en duro); se mitiga con los goldens fila-a-fila + `<f>` byte-identica + DetRetri-D vs R10 y con el re-congelamiento T0 ante forma nueva. **Evidencia:** `T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback` verde sin tocar el test, y set-diff vacio (39=39) en `p45-h2b.trx`.

**Fail-fast:** fila cuya (firma, seccion-contrato) no tiene slot en el contrato T0 y NO esta en la lista de cero documentado -> `ERR-SIN-SLOT` nombrando ASE + firma A|B|C|D|E + rotulo de empresa + canal + concepto + seccion-contrato esperada + ocurrencia + celda candidata. Prohibido resolver por proximidad, orden fisico o heuristica posicional (E1-1 Plan 41). Fila sin contraparte en fuente -> slot queda en cero predefinido (no se escribe, no es error).

### 2.2 Regla sin-C + cero documentado (E-44.1, derivada en H0)

- **Regla sin-C:** las filas sin empresa en C se colocan por (firma + seccion-contrato-por-ASE). La seccion-contrato la fija la tabla H0 por ASE (forma de seccion, no etiqueta de frontera sola): p. ej. la `Aplic#1` fuente ASE5 (VS/VI con ENERBIT) alinea con la `Aplic#1` oficial (630-647), no con la `Oportuno#1` oficial; el bloque `Componente/Subsidio` fuente alinea con el `Oportuno#2` oficial (649-659); `Aplic#2` y cierre `Subs/Cont`/`Total` alinean exacto. **Fallback por orden fisico prohibido.**
- **Cero documentado:** las filas sin slot de la lista H0 (Intereses de 1.a ASE5: 402,40 y 7.697,23 + desglose, + las que liste el T0 por ASE con valores diferidos citados) = cero documentado + reporte al dueno (texto T0 padre S9, extendido por ASE en H0). Sin `ERR-SIN-SLOT`, sin simular. Los Totales ya llevan el dinero (O 50.323,46 = 49.921,06 + 402,40; T 314.490 = 306.792,77 + 7.697,23): solo se difiere el desglose.
- **`ERR-SIN-SLOT` residual:** queda SOLO para fila con (firma, seccion-contrato) sin slot FUERA de la lista documentada, con fail-fast que nombra ASE+cuadro+fila.
- **F654/F661:** la regla DEBE taparlo (artefacto de escritura probado T0 S5). El T0 propio lo re-verifica (lector emite una vez por fila; escritor por seccion-contrato no duplica `Subsidio` y `D=H/Total` en 654/661).
- **I via contrato explicito; F como rama descartada:** E-44.1 punto 4. Intacto del T0 padre: T solo T, `U->Z`, `Exonera-*` intactas, auditoria pendiente, Via A/B en F2 del padre, Q1 y otras pestanas fuera.

### 2.3 Colocacion dinamica multicolumna values-only (heredada del padre, intacta)

- Escritura por encabezado F..Z (padre F1 DONE): cada encabezado fuente -> mismo encabezado plantilla; `Rec.Intereses` U -> Z; Exonera-* intactas; `Intereses Cte` solo T (Y intacta, no espejo).
- Values-only: solo `<v>` numerico con decimal exacto de la fuente (sin redondeo; redondeo a entero exclusivo de DetRetri, Plan 29 Unidad D). Guard anti-formula intacto: celda destino con `<f>` -> `ERR-PLANTILLA`. Cero `InsertRow`/`DeleteRow`, cero `new CellFormula` fuera de `EscribirFormulaPreservando`, cero Reanclar en este path. Slots no consumidos quedan en cero.
- `RangoDatosAse` intacto; F682/F702 (EXTEMPORANEO ASE5, literales 0) siguen FUERA del rango salvo arbitrio UAESP (residual 24.329,38: F2 del padre; sin arbitrio no se incluyen).

### 2.4 Que filas/columnas NO se tocan

- Anclas y visibles (`F75/F254/F415/F568/F680`, `F80..F705`, `F700`, `F682/F702`, `CONSOLIDADO!D9:D13/D51/D108/D109`, `REMUNERACION_*!D9:D13/D51/D108`, `VALIDACION_*`, `DetRetri/DetValiRetri`, `INTERVENTORIA`, `N3/G7/K7/D6/D7`, Phase 0 allow-list de 68 celdas): no se tocan en H1 (las recompone/evalua su propio dueno).
- Columnas Exonera-* (U/V/W/X) sin contraparte: intactas. Columna Y (segundo `Intereses Cte`): intacta.
- Fuera de los bloques espejo R1 (`RangoDatosAse`): prohibicion insert/delete intacta.
- Q1 (codigo, tests, geometria, writer): no se toca. `DivergenciasDeclaradasCatalogo`: no se toca.

---

## 3. SPEC (criterios por tarea + gates)

### S0 - Gates transversales (toda T)

- G-WORKBOOK: todo gate de negocio lee el **workbook generado** (texto-`<f>` + literales; prohibido `<v>` stale, doctrina Plan 28/29) y lo contrasta contra el dominio por firma +-0.5.
- G-F-BYTE: `<f>` byte-identica en julio-identidad (recomposicion Plan 31 + shared-formula Plan 33: `EscribirFormulaPreservando`, invariante `si` workbook-wide).
- G-DETRITRI-R10: DetRetri calculado bottom-up (`ROUND(D104:D108,0)`) vs R10 (`DetRetri{AAAAMMQ}` D9:D13 + D14) +-0.5 post-redondeo en AMBAS quincenas; divergencia -> fail-fast con periodo + archivo + ambos valores. Intacto (insumo para F2 del padre, no re-decision aqui).
- G-AUDITORIA (al llegar): la auditoria manual es oraculo de paridad; sin auditoria no hay verde por ausencia.
- G-SHA: oficial SHA256 intacto antes/despues (`3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).
- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` 0 warnings / 0 errores; CRLF sin mezclar; NUNCA `git commit`/`push`; rutas explicitas; regens a temp FRESCA con el OFICIAL como unica entrada; prohibido `Docs/Insumos` raiz en lo nuevo; tests NO se debilitan (0 `Skip`, 0 `Assert.True(true)`, 0 `NotImplementedException`).
- G-Q1-INTOCADO: set-diff Q1 sin verdes nuevos ni rojos nuevos sin autorizacion explicita separada; `SeccionEspejoR1Tests.T1_Ase2Q1` sigue documentando el bloqueo Q1; grep: cero cambios de comportamiento Q1 en produccion.
- G-D-H: cero literales/ramas por periodo o empresa en produccion (grep de cierre).

### S1 - Criterios H0 (T0 multi-periodo propio, read-only)

- S1.1: fixture read-only cita bytes del oficial (rangos por ASE, encabezados F..Z, filas 6-17 / 110-121 / 289-302 / 450-461 / 603-616, slots 612/628/647/659/673 + seccion por frontera, `F700` 3 terminos, `F682/F702` literales 0, `F680` terminos) + del manual legacy (O 402,40 / T 7.697,23 ASE5) + de las fuentes de los 3 periodos (julio-2 filas 8/11 O/T con desglose; agosto ASE5 3 secciones `Mes`; Q1-ASE2 `D=5` en 1.a y ausencia de `Aplicacion`, read-only).
- S1.2: fronteras y formas por ASE x periodo publicadas (tabla: ASE x periodo -> secuencia de fronteras con filas citadas; inversion ASE5 julio-2; 3 `Mes` agosto ASE5; Q1 sin `Aplicacion`).
- S1.3: matriz fuente<->oficial completa por ASE x periodo (cada celda con cita fila-fuente + celda-slot + seccion-contrato); `Aplic#2` y cierre identicos ASE5 citados; `Aplic#1` fuente<->oficial (con ENERBIT) citada; bloque `Componente/Subsidio` citado.
- S1.4: regla sin-C derivada sobre bytes reales de los 3 periodos (tabla de decision por forma, geometria, cero ramas); lista documentada de ceros con valores diferidos citados (402,40 / 7.697,23 + desglose + las que liste por ASE); aislamiento F654/F661 re-verificado (lector una emision por fila; escritor por seccion-contrato lo tapa).
- S1.5: formas Q1 citadas sin resolver (Q1-ASE2 `D=5`, Q1-ASE5 ENERBIT en Oportuno, ausencia de `Aplicacion`); constancia explicita de que Q1 queda para su propio contrato.

### S2 - Criterios H1 (implementacion regla sin-C + cero documentado)

- S2.1: filas sin C colocadas por (firma + seccion-contrato-por-ASE) sin orden fisico; regresion de las 11 colocaciones medidas ASE5 corregidas por forma (F604/F610/F616/F620/F623/F626/F631/F635/F639/F613/F648 + 674 en 0 segun matriz H0; valores exactos de fuente por encabezado F..Z).
- S2.2: filas con empresa en C intactas (F612=417.519,62, F628=24.329,38, F647 con L647, F673 y hermanas por ASE; muestra + barrido 5 ASE).
- S2.3: filas de la lista documentada en cero documentado sin `ERR-SIN-SLOT` (Intereses 1.a ASE5 402,40 / 7.697,23 + desglose + lista H0); Totales verificados (O/T citados); fuente sin slot fuera de lista = `ERR-SIN-SLOT` con ASE+firma+empresa+canal+concepto+seccion-contrato+ocurrencia.
- S2.4: F654/F661 sin duplicado (cada fila `Subsidio` / `D=H/Total` en su slot de seccion-contrato; `-187.855,98` una vez donde corresponde, cero donde no).
- S2.5: `<f>` intacta en todo el diff H1 (grep: cero `new CellFormula` fuera de `EscribirFormulaPreservando`; cero `InsertRow`/`DeleteRow`); D-H (cero literales/ramas por periodo o empresa); tolerancia +-0.5 solo en comparacion, nunca en escritura (decimal exacto); Y y Exonera-* intactas.

### S3 - Criterios H2 (regresion multi-periodo)

- S3.1: corrida Q2 (2026072) desde el OFICIAL a temp FRESCA (plantilla CERRADA): F680=12.033.011.685,71; D34/F34/O25 a 417.519,62 exactos salvo veredicto F2 del padre; C15/D25 TRUE; DetRetri-D vs R10 intacto; agosto-2026082 con sus goldens (ASE5 3 `Mes` absorbidos).
- S3.2: suite completa con trx + set-diff vs linea base **39 Failed / 888 Passed / 927 Total** (mismos 39; 8 de Q1 declarados intactos; cero verdes Q1 nuevos); 0 warnings; SHA256 intacto.
- S3.3: tabla observado-vs-esperado por celda (O25/D34/F34/C15/D25 + DetRetri-D vs R10 por ASE) entregada como insumo a F2 del padre; Via A/B NO se re-decide aqui.

### S4 - Criterios H3 (docs)

- S4.1: `project-context.md` con doctrina contrato PERIOD-FREE + regla sin-C + lista de ceros + formas multi-periodo + Q1 citado-sin-resolver.
- S4.2: nota al manual HU-17 (instructivo Capa B): que auditar fila-a-fila por (fuente, canal, concepto, seccion-contrato) y donde mirar el reporte de divergencias declaradas + reporte de ceros al dueno.
- S4.3: reporte al dueno del hueco `Vlr Intereses` (texto listo del padre S9 extendido por ASE, sin enviar plantilla).

---

## 4. TASKS

### H0 - T0 multi-periodo propio con fixtures (bloqueante; cero codigo de produccion)

1. H0.1 Inventario BCL (ZIP+XML, sin Excel/COM/OpenXML): encabezados F..Z del oficial por bloque ASE + encabezados F..U de fuentes R1 (julio-2 5 ASE, agosto-2026082 5 ASE, Q1-2026071 muestra read-only) + fronteras por ASE x periodo (filas citadas: julio-2 5/5/5/5/4 con inversion ASE5; agosto ASE5 3 `Mes`; Q1 sin `Aplicacion` + Q1-ASE2 `D=5`) + filas fuente julio-2 8/11 (O 402,40 / T 7.697,23 con desglose) + filas sin C medidas + `F654`/`F661` + `F612`/`F613`/`F680`/`F700`/`F682`/`F702` + `F628` huerfano (grep consumidores todas las hojas, maestros shared incluidos). Rutas explicitas (oficial + `Docs/Prueba Julio-2/Insumos/...` + insumos agosto + `Docs/Prueba Julio-1/Insumos/...` read-only). Prohibido `Docs/Insumos` raiz. SHA256 antes/despues.
2. H0.2 Matriz fuente<->oficial completa por ASE x periodo + regla sin-C derivada sobre bytes reales de los 3 periodos (tabla por forma, geometria, cero ramas) + lista documentada de ceros con valores diferidos citados + re-verificacion F654/F661 (lector vs escritor) + formas Q1 citadas sin resolver. Fixture: `Remuneracion.IntegrationTests/ContratoSeccionPeriodFreeT0Tests.cs` (solo lectura + asserts de bytes; no toca produccion; Q1 solo lectura).
3. H0.3 Congelar ejemplos del Ingeniero + redactar reporte al dueno (hueco `Vlr Intereses` por ASE, texto listo sin enviar) + constancia Q1 (citado-sin-resolver, requiere contrato propio con autorizacion separada).
4. Gate H0: fixture verde + doc `Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md` con tablas citadas + regla + lista de ceros + formas Q1 + veredicto F654/F661. Sin esto no se abre H1. Si la regla no absorbe las 3 formas sin ramas, STOP declarado (no se implementa).

### H1 - Implementacion regla sin-C + cero documentado (contract-first)

1. H1.1 Core puro: contrato de seccion-por-ASE (predicados puros, sin I/O, sin periodo; catalogo abierto; D-H) + `ERR-SIN-SLOT` residual enriquecido (ASE+firma+empresa+canal+concepto+seccion-contrato+ocurrencia). `EsDatoEmpresa` y `EsAplicacionPorFrontera` INTACTOS.
2. H1.2 Consumo en `EscritorValoresSlotR1`: grupos por (firma + seccion-contrato), cursor por grupo, cero documentado (lista H0; sin `ERR-SIN-SLOT`, sin simular), guard anti-formula intacto, decimal exacto, Y y Exonera-* intactas. `RangoDatosAse` y geometria intactos.
3. H1.3 Cableado (ProcesadorPeriodo/ProcesadorRemuneracion + mapa leaf R1-Q2, Plan 25 intacto): el agregado `totOpt`/`extemp` sigue mandando en CONSOLIDADO; el interior multicolumna alimenta las anclas por empresa. Sin ramas por periodo/empresa.
4. H1.4 (review F1 SPT-1 plegada): encabezado de fuente sin columna destino en el template -> `Log.Warning` accionable (ASE+fila+encabezado) o `ERR-FUENTE`; nunca descarte silencioso. (Hoy inocuo por T0 §1.2 del padre; red contra encabezados futuros de fuente.)
5. H1.5 (review F1 SPT-2 plegada): `BloqueEspejoAseInputs.TieneColumnaEspeciales` huérfano en producción (la escritura de L ahora se dispara por clave con valor): documentar su rol fixture-only o retirarlo con sus consumidores, sin cambiar semántica.
6. Gate H1: tests focalizados (S2.1/S2.2/S2.3/S2.4) verdes + build 0 warnings + greps (0 `<f>` fuera de Phase 0/`EscribirFormulaPreservando`, 0 Insert/Delete, 0 literales por periodo/empresa, Q1 intacto).

### H2 - Revision + 0 warnings + set-diff + corridas frescas (calidad senior)

1. H2.1 Code-review independiente (spec compliance + calidad): trazabilidad cada cambio -> requisito S2; ningun test debilitado; CRLF sin mezclar; 0 recursos UI inventados (no hay UI).
2. H2.2 Suite completa con trx + set-diff vs linea base **39 Failed / 888 Passed / 927 Total** (cero rojos nuevos; 8 de Q1 declarados intactos; cero verdes Q1 nuevos). Comando: `dotnet test Remuneracion.WinForms/Remuneracion.WinForms.slnx --logger "trx"` con `--filter` focalizados primero.
3. H2.3 Corridas frescas: regen Q2 (2026072) y agosto (2026082) desde el OFICIAL a temp FRESCA (rutas explicitas; plantilla CERRADA) + goldens fila-a-fila + `<f>` byte-identica (julio-identidad) + DetRetri-D vs R10 por ASE + tabla observado-vs-esperado para F2 del padre (Via A/B NO se re-decide).
4. Gate H2: review APROBADO (o APROBADO CON OBSERVACIONES cerradas) + set-diff vacio (salvo veredicto documentado) + SHA256 intacto + G-Q1-INTOCADO verde. Sin esto no se abre H3.

### H3 - Docs (project-context.md + manual + reporte)

1. H3.1 Actualizar `.opencode/project-context.md` (doctrina Plan 45: contrato PERIOD-FREE + regla sin-C + lista de ceros + formas multi-periodo + Q1 citado-sin-resolver + F654/F661 tapado).
2. H3.2 Nota al manual de HU-17 (instructivo Capa B): auditoria fila-a-fila por (fuente, canal, concepto, seccion-contrato) y donde mirar reportes (divergencias declaradas + ceros al dueno).
3. H3.3 Reporte al dueno del hueco `Vlr Intereses` (texto listo H0.3; NO tocar la plantilla; NO enviar sin aprobacion).
4. Gate H3: diff de docs citado en el resumen de cierre; cero codigo de produccion en H3.

---

## 5. Invariantes vigentes (todos heredados, sin excepcion)

Values-only; jamas `<f>` fuera de Phase 0 (allow-list 68 celdas: 16 Plan 39 + 36 recaudo + 16 F-NEW-1); cero `InsertRow`/`DeleteRow`; preservacion shared-formula (Plan 33); D-H (cero literales/ramas por periodo o empresa); tolerancia +-0.5 solo en comparacion; CRLF sin mezclar; NUNCA `git commit`/`push`; oficial SHA256 intacto antes/despues; quincena por dominio (`Periodo.NumeroQuincena`), nunca por contenido; tests NO se debilitan; declarada = ultima puerta con doble huella y fail-closed (Plan 42); prohibido fallback posicional (Plan 41 E1-1) y prohibido sellar/apagar/re-etiquetar (Plan 40 E1-1..E1-5); Q1 ni se toca ni se reverdece (RESTRICCION DURA 0.2); `<v>` nunca es oraculo; Enmienda E-44.1 y puertas del padre (auditoria pendiente, Via A/B en F2) intactas y referenciadas, no reabiertas.

Comandos:

```bash
dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx
dotnet test Remuneracion.WinForms/Remuneracion.WinForms.slnx --filter "<focalizados>" --logger "trx"
dotnet test Remuneracion.WinForms/Remuneracion.WinForms.slnx --logger "trx"
```

Rutas explicitas en todo lo nuevo; prohibido `Docs/Insumos` raiz; regens siempre a temp FRESCA con el OFICIAL como unica entrada y la plantilla CERRADA en Excel.

---

## 6. Architecture Validation Certificate

- **SOLID:** SRP (contrato de seccion-contrato puro en Core / lectura por encabezado en reader / pegado values-only en writer / clasificacion en `ValidadorBasico` + catalogo de datos + lista de ceros como DATOS del T0); OCP (nueva forma/periodo = fila citada en T0 nuevo, cero ramas nuevas; periodo nuevo = cero cambios por D-H y PERIOD-FREE); LSP/DIP (writer depende de `BloqueEspejoAseInputs` + predicados + contrato, no de concretos de periodo); ISP (contratos `IRecaudoReader`/`IPlantillaWriter`/`IDetRetriR10Reader`/`IReporteDivergenciasDeclaradas` intactos).
- **Best Practices:** contract-first (H0 congela bytes multi-periodo antes de H1); fail-fast accionable (`ERR-SIN-SLOT` residual / `ERR-PLANTILLA` con ASE+firma+seccion-contrato+celda); gates que leen el workbook (texto-`<f>` + literales, nunca `<v>`); goldens +-0.5 + `<f>` byte-identica + DetRetri-D vs R10; set-diff contra linea base + G-Q1-INTOCADO; SHA256 del oficial; CRLF; sin commits.
- **Performance:** resolucion por encabezado una vez por hoja (mapa encabezado->indice cacheado por corrida); contrato por ASE resuelto una vez por bloque (tabla citada, sin re-lectura por fila); agrupacion por (firma, seccion-contrato) con cursor por grupo (O(filas + slots)); lectura acotada en H0 (BCL ZIP+XML); sin abrir workbooks en preflight; regens a temp fresca sin copiar de mas.

---

## 7. Clarification Gate (que necesita aprobacion del Ingeniero)

1. **Contrato PERIOD-FREE + regla sin-C + cero documentado (H0/H1):** ratificar la tabla fuente<->oficial por ASE x periodo, la regla derivada sobre bytes reales de los 3 periodos y la lista documentada de ceros con valores diferidos citados antes de abrir H1.
2. **Q1 citado-sin-resolver:** ratificar que Q1 queda fuera de codigo en este plan (solo lectura read-only en H0) y que cualquier cambio Q1 exige autorizacion explicita separada.
3. **Reporte del hueco `Vlr Intereses` al dueno:** ratificar texto (H0.3/H3.3); NO tocar la plantilla; NO enviar sin aprobacion.
4. **Puerta H0->H1:** sin H0 verde (fixture + doc T0 propio) no se implementa. Si H0 declara STOP (regla no absorbe las 3 formas sin ramas), no se abre H1.
5. **F2 del padre:** Via A/B y auditoria siguen en el padre con numeros frescos post-hijo; este plan solo entrega la tabla observado-vs-esperado como insumo, no re-decide.

Sin estas aprobaciones el plan no avanza de H0 a H1 (1-2, 4) ni cierra H2 (toda la regresion) ni H3 (3, 5).

---

## 8. Cierre (2026-10-16) y follow-ups declarados

**Estado del gate (los 5 puntos del §7):** 1 ratificado 2026-10-16 (P2 + cero documentado, E-44.1) · 2 ratificado (Q1 intocado; test de Q1 verde sin cambios) · 3 texto listo, **no enviado** (sin tocar la plantilla) · 4 cumplido (H0 verde → H1 abierto) · 5 trasladado al padre (F2).

**Evidencia de cierre:** build 0 warnings/0 errores · focalizados 23/23 (incl. `T1_Ase2Q1…` verde sin tocar el test) · suite `p45-h2b.trx` = 968/929/39/0 con **set-diff VACIO** vs `o43-final-failed.txt` (39=39, cero nuevos, cero desaparecidos) · oficial SHA256 intacto · review H2.1 APROBADO-CON-OBSERVACIONES con **O-1 formalizada en §2.1** · CRLF limpio · sin commits.

**Follow-ups declarados (NO de este plan; salen a plan propio si se priorizan):**
- **O-2:** los dos call-sites de produccion (`OpenXmlPlantillaWriter.cs:297/437`) descartan el resultado de `Escribir`; falta senal agregada (nº de ceros documentados / encabezados sin destino por corrida) y test del warning de H1.4.
- **O-3:** el test del `ERR-SIN-SLOT` enriquecido no congela firma/canal/concepto/seccion/ocurrencia (la implementacion ya los nombra).
- **Padre 44 / F2:** Via A/B sobre OCCIDENTE-ASE5 con numeros frescos + **auditoria de la propietaria** (oraculo pendiente, no verde por ausencia) + reporte del hueco `Vlr Intereses` al dueno.
- **Otras pestanas:** misma lente (fuente/canal/concepto/componentes), plan propio.
