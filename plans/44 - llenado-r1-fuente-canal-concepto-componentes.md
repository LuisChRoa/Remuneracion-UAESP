---
titulo: Plan 44 - Llenado R1 fuente-canal-concepto-componentes
estado: F1-MULTICOLUMNA DONE y verificado por el orquestador 2026-10-16 (build 0w, 22/22, anchor-neutral, set-diff vacío); regla sin-C + cero documentado RESUELTA EN PLAN HIJO 45 (CERRADO 2026-10-16) y CORREGIDA POR EL PLAN 46 (CERRADO 2026-10-17: sección por CIERRE + canal efectivo; la regla de sección del 45 y su cero documentado quedan SUPERSEDIDOS — ver nota en plans/45); **F2 — VÍA A/B DISUELTA por el cierre del 46**: el Δ de OCCIDENTE ASE5 era un defecto de MAPPING nuestro (F612→0; F700 = 304.989.620,81 = medición del administrativo; VALIDACION_TOTAL O25/D34/F34 TRUE, Δ 0), NO una regla UAESP pendiente; **F3/paridad-auditoría PENDIENTES: la auditoría de la propietaria sigue sin entregarse y NO se declara paridad sin ella**
fecha: 2026-10-16
numero: 44
planes_ocupados: 01..43 ocupados; este es el 44
gate_entrada:
  alcance: R1 completo + re-verificacion de autocontroles/divergencia + nuevas pruebas Julio-2
  oraculo: auditoria manual de la propietaria PENDIENTE de entrega (puerta explicita, no verde por ausencia)
  q1: frente propio (8 rojos declarados, no tocar)
  otras_pestanas: follow-up explicito (misma lente, plan hijo)
---

# Plan 44 - Llenado R1 fuente-canal-concepto-componentes (fuente x canal x concepto x componentes)

> SDD completo PROPOSE -> DESIGN -> SPEC -> TASKS. Fuente unica de verdad para la implementacion una vez aprobado ESTE documento.
> Lectura obligatoria previa (rector irrenunciable): `README.md`, `.opencode/project-context.md` (doctrinas Planes 21-43), `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md`, mas evidencia `Docs/T0-Plan41-MiniT0-OccidenteAse5-SaldosNota-TablaJ.md` p.1, `Docs/T0-Plan42-Taxonomia-DivergenciasDeclaradas.md` p.2, `Docs/T0-Residual-24329-OccidenteAse5.md`, `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs`, `Remuneracion.Core/Models/R1FirmaInterior.cs`, `Remuneracion.Core/Models/DivergenciasDeclaradasCatalogo.cs`.
> Entorno: Windows, repo con espacios en la ruta; la plantilla oficial debe estar CERRADA en Excel al correr tests (FileShare). Oficial SHA256: `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` (verificar intacto antes/despu\u00e9s).

> **Nota F2 (2026-10-17) — Vía A/B DISUELTA.** El cierre del Plan 46 demostró que el Δ de 417.519,62
> de OCCIDENTE ASE5 era un **defecto de mapeo nuestro** (colocación del interior R1), no una regla de
> negocio UAESP pendiente: con la sección por frontera de CIERRE + canal efectivo, `F612` = 0,
> `F700` = 304.989.620,81 (igual a la medición del administrativo) y `VALIDACION_TOTAL!O25/D34/F34`
> sellan TRUE (Δ 0). Consecuencias: (a) la entrada `OCCIDENTE-ASE5` del catálogo de divergencias
> declaradas quedó RETIRADA (con fecha/motivo/evidencia); (b) NO hay decisión Vía A/B que tomar —
> queda disuelta; (c) la **auditoría de la propietaria sigue PENDIENTE** y NO se declara paridad sin
> ella. Evidencia: `plans/46` §9, `p46-cierre.trx` (38/960/998, set-diff 0 altas · 1 baja).

## Enmienda E-44.1 — P2 + cero documentado (ratificada por el Ingeniero 2026-10-16; reabre F1)

**Decisión:** P2 (contrato de sección explícito por ASE, citado en bytes; geometría, cero ramas por período) + filas sin slot como **cero documentado** + reporte al dueño. Levanta el STOP de F0 (§10 del T0) y reabre F1→F2→F3 bajo esta enmienda.

**Qué cambia respecto a §2.1/§2.2:**
1. Regla sin-C: las filas sin empresa en C se colocan por (firma + **sección-contrato**): la sección de la fila fuente la resuelve el contrato explícito por ASE (tabla fuente↔oficial citada en el T0), no el (tipo, ordinal) genérico. **Fallback por orden físico prohibido** (era el defecto).
2. Filas sin slot (Intereses de 1.ª ASE5: 402,40 y 7.697,23 + desglose, + las que liste el T0 por ASE) = **cero documentado** + reporte al dueño (texto T0 §9). Sin `ERR-SIN-SLOT`, sin simular. Los Totales ya llevan el dinero (O 50.323,46 = 49.921,06 + 402,40; T 314.490 = 306.792,77 + 7.697,23): solo se difiere el desglose.
3. `ERR-SIN-SLOT` queda SOLO para fila con (firma, sección-contrato) sin slot **fuera** de la lista documentada, con fail-fast que nombra ASE+cuadro+fila.
4. I se adopta vía contrato explícito (punto 1); F como rama queda descartada. F654/F661 **debe** taparlo la regla (artefacto de escritura probado).
5. Intacto del T0: T/Y solo T · `U→Z` · `Exonera-*` intactas · auditoría pendiente (paridad al llegar) · Vía A/B solo en F2 con números frescos · Q1 y otras pestañas fuera.
6. Higiene: línea base vigente **39 Failed / 888 Passed / 927 Total** (mismos 39 fallos; el plan citaba 912/873 del Plan 42 T5). Criterio vinculante: mismos 39 fallos + 12 del fixture F0 verdes.

---

## 0. Contexto verificado en bytes (adoptado, no re-diagnosticado; cada gate lo re-verifica por test)

### 0.1 Modelo correcto (dado por la propietaria de la plantilla estandar, verificado en bytes)

- **Componentes = encabezados de columnas.** Estandar: F=Total, G=TDF, H=TTL, I=TVIAT, J=Aprovechamiento, K=CCSA Prest.Aprov., L=SERVICIO ESPECIALES, M=TCS, N=TLU, O=TBL, P=TRT, Q=CCSA No Aprov., R=Ajuste Decena, S=Ajuste Recargo, T/Y=Intereses Cte, U/V/W/X=Exonera-*, Z=Rec.Intereses. La fuente trae F..U con T=Intereses Cte, U=Rec.Intereses.
- **Mapeo de columnas POR ENCABEZADO**, nunca por posicion: Rec.Intereses U(fuente) -> Z(plantilla); Exonera-* (U/V/W/X plantilla) sin contraparte en fuente = se dejan intactas (cero predefinido del template).
- **Fuentes de recaudo = columna C** (ENEL, OCCIDENTE, NUEVO ESQUEMA, ENERBIT; mas RECIPROCIDAD legado por sinonimia Plan 34, CIUDAD LIMPIA-ACUEDUCTO donde aplique).
- **Canales = columna D** (H, O, T, E, 1, 2, 5).
- **Conceptos = columna E** (Vlr Servicio, Vlr Intereses, Total; mas Subsidio, Mes, Aplicacion nuevos x reversion como fronteras/roles).
- **Llenado DINAMICO desde fuentes:** lo que llega se coloca donde corresponde; lo ausente queda en el cero predefinido. Ejemplo correcto: Vlr Intereses del canal H en 0 y AFaseo vacio de Area Limpia son CORRECTOS asi (ausencia de fuente, no hueco).

### 0.2 Tres desvios medidos de la app (julio-2, corrida `Docs/Prueba Julio-2/Resultado/6` vs fuente `5-Area Limpia Recaudoporcomponente` y oficial)

(a) **Solo F (Total) y L (Especiales) por fila; G..AP quedan en 0** aunque la fuente trae el desglose. Ejemplo fila fuente OCCIDENTE 1.a: Total 417.519,62, TDF 26.657,43, TTL 14.138,02, TVIAT 7.723,27, Aprov 93.911,73, CCSA 2.290,29, TCS 20.089,34, TLU 28.092,36, etc. La app hoy solo pega F y L (`EscribirValor` sobre `CeldaF`/`CeldaL`).

(b) **Filas sin empresa en C casan SOLO POR ORDEN fisico** (`EsDatoEmpresa` exige C no vacia; si un lado es null la seccion no veta; `EncajaEnSlot` compara A/B/D/E + C por sinonimia, y para filas no-dato la seccion es null). Medido en ASE5 (bloque 603-676): F604=12.823.196.468,26 (E-Aplic#1 en slot E-1.a), F610=12.815.008.425,00 (Servicio ENEL-Aplic#1 en slot T-1.a), F616=402,40 (Intereses O-1.a en slot E-Aplic#1), F620=7.697,23 (Intereses T-1.a en slot H-Aplic#1), F623=8.188.043,26, F626=90.800,20, F631=203.478,11, F635=1.803.780,95, F639=1.037.886,13, F613=13.084.233.937,44 (Mes-Aplic#1 en slot Mes-1.a), F648=-1.036.320.522,40 (Mes-Aplic#2 en slot Mes-Aplic#1; 674 queda en 0). Invisible a anclas/validacion (solo leen filas Total/empresa/Mes agregadas). Micro-punto abierto: **F654=-187.855,98 duplica a F661** (mecanismo sin aislar: emision del lector vs escritor; lo verifica el T0, no se supone).

(c) **Primer bloque de la estandar trae CERO filas `Vlr Intereses` en los 5 ASE** (filas 6-17, 110-121, 289-302, 450-461, 603-616 verificadas); la fuente SI trae intereses de 1.a seccion (filas 8: O 402,40 y 11: T 7.697,23, con desglose) y la plantilla anterior SI los registraba en ASE5 (O 402,40, T 7.697,23). Hueco de PLANTILLA a reportar al dueno (**NO tocar la plantilla**; SHA256 oficial intacto). OJO no sobregeneralizar: en ASE1 la anterior tampoco traia Intereses de 1.a; el T0 mapea por ASE/cuadro. Las filas con empresa en C (612, 633, 647, 663, 673, etc.) quedaron CORRECTAS.

### 0.3 Contexto de validacion (re-verificar con numeros frescos, NO prejuzgar)

- F612 solo alimenta a F700 en todo el libro (scan exhaustivo); F613 solo a F680; F680=12.033.011.685,71 coincide con el golden/manual (el total del ASE esta bien).
- D34/F34/O25 difieren en 417.519,62 exactos = F612-L612; entrada declarada OCCIDENTE-ASE5 vigente con dueno UAESP (`DivergenciasDeclaradasCatalogo`); residual 24.329,38 documentado (`Docs/T0-Residual-24329-OccidenteAse5.md`: F628 huerfano, F682/F702 literales 0, ASE1-4 con formula EXTEMPORANEO que ASE5 no trae).
- Tras el re-llenado correcto, los autocontroles pueden moverse: el plan re-verifica y re-decide Via A/B con evidencia fresca, sin presuponerla.

---

## 1. PROPOSE

### 1.1 Intent

Cerrar el frente R1-interior por **contenido** (no solo por geometria): cada fila fuente `(C=fuente, D=canal, E=concepto, seccion-por-frontera)` coloca **todos sus componentes** en el slot `(firma + seccion)` que le corresponde, por encabezado de columna, values-only, sin insert/delete, sin tocar `<f>`. Re-verificar con numeros frescos los autocontroles (W-1 C15/D25/O25/D34/F34) y la divergencia declarada OCCIDENTE-ASE5, y dejar pruebas Julio-2 que demuestren el cierre fila-a-fila contra la fuente.

### 1.2 In-scope EXPLICITO por frente

| Frente | Alcance | Criterio de salida |
|---|---|---|
| **F0 - T0 evidencia read-only** | Congelar geometria + ejemplos del Ingeniero (filas 6-17 ASE1, 619-628 ASE5) + matriz fuente(fuente,canal,concepto,seccion)->slot(F..Z) por ASE + veredicto fail-fast vs fallback para filas sin C + aislamiento del micro-punto F654/F661 + paridad pendiente contra auditoria (puerta explicita). Cero codigo de produccion. | Fixture read-only verde que cita bytes; decisiones F1/F2 desbloqueadas o declaradas STOP. |
| **F1 - Llenado R1 multicolumna** | Lector R1 por encabezado (F..U fuente) + escritor slot multicolumna values-only (F..Z plantilla, Rec.Intereses U->Z, Exonera-* intactas) con clave (C-por-sinonimia, D, E, seccion-por-frontera). Incluye regla de filas sin C decidida por F0 y correccion del desvio (a) y (b). | Slots F/L/G..Z por fila coinciden con fuente por encabezado fila-a-fila en Julio-2; filas sin contraparte quedan en cero predefinido sin `ERR-SIN-SLOT` espurio. |
| **F2 - Re-verificacion autocontroles/divergencia** | Re-correr W-1 (C15/D25/O25/D34/F34) + DetRetri-D vs R10 + `DivergenciasDeclaradasCatalogo` con numeros frescos tras F1; re-decidir Via A/B (mantener declarada / ampliar / retirar por condicion de retiro) con evidencia, sin presuponerla. | Tabla observado-vs-esperado por celda + veredicto por celda + reporte durable si aplica; ningun gate apagado/sellado/re-etiquetado. |
| **F3 - Nuevas pruebas Julio-2** | Corridas frescas Julio-2 (2026072) desde el OFICIAL a temp FRESCA + goldens fila-a-fila + `<f>` byte-identica + set-diff de suite. | T3 verde con evidencia citada; linea base de suite respetada. |

### 1.3 Non-goals / out (vinculante)

- Otras pestanas (R2-detalle, R4-detalle, SALDOS POR NOTA, RETRIBUCION NEGATIVA, AJUSTES-SF-T, BCE, banco, conciliaciones, INTERVENTORIA salvo lo ya doctrinal, AFaseo-consumidores): **follow-up explicito** (misma lente, plan hijo). No se tocan en este plan.

> **FOLLOW-UP §1.3 ATENDIDO (2026-10-17).** El plan hijo es `plans/47 - revision-demas-pestanas-lente-r1.md`
> (APROBADO 2026-10-17). Su T0 read-only (`Docs/T0-Plan47-Revision-Pestanas-LenteR1.md`) probó por bytes
> la tabla dentro/fuera por pestaña y el desvío de `Rem. Anticipos R2` / `Reversion Pagos R4`; con la
> autorización del Ingeniero se ejecutó H1 (canal efectivo R46-1 análogo + canal compuesto por segundo
> cierre, ceros documentados E-44.1 y `Otros` por encabezado) y H2 (corrida fresca + set-diff). Las
> pestañas declaradas **fuera** quedaron **intactas con evidencia** (SALDOS POR NOTA, `Recaudo *`, BCE,
> CONSOLIDADO, AJUSTES, VALIDACION_*, GERENTES_*, REMUNERACION_*, DetRetri*: 0 `values_changed`).
- Q1 (2026071): **frente propio** (8 rojos declarados por contrato de seccion faltante, Plan 41 E1-4/E1-6). No tocar, no reintroducir fallback, no reverdecer por apano.
- Plantilla: **NO se edita** (ni el oficial ni la copia de trabajo fuera de Phase 0 allow-list). El hueco `Vlr Intereses` del primer bloque se **reporta al dueno**, no se parchea con filas.
- Auditoria manual de la propietaria: **pendiente de entrega**; la prueba de paridad queda PENDIENTE (puerta explicita, no verde por ausencia).

### 1.4 Riesgos con mitigacion

| ID | Riesgo | Mitigacion |
|---|---|---|
| R-ATRIBUCION cruzado | Atribuir a colocacion lo que es plantilla (o viceversa): p. ej. F628 huerfano vs error de slot; F654/F661 duplicado (lector vs escritor). | T0 aisla por bytes: grep exhaustivo de consumidores `<f>` (maestros shared incluidos) + traza lector (filas emitidas) vs escritor (slots consumidos) por (firma,seccion,ocurrencia). Veredicto citado por celda; sin veredicto no hay fix. |
| R-AUDITORIA-pendiente | La auditoria manual aun no llega; presionar un verde sin oraculo. | Puerta explicita: T0 congela geometria + ejemplos; SPEC deja el gate workbook-vs-dominio con la auditoria como oraculo **al llegar**; mientras tanto oraculos vigentes (fuente por encabezado, goldens, R10, `<f>` byte-identica). Prohibido declarar exito por ausencia. |
| R-PLANTILLA-nueva-version | El negocio entrega una version nueva del oficial (p. ej. con filas `Vlr Intereses` en el primer bloque). | Contrato por etiquetas, no por direcciones: `ContratoSlotR1T0Tests` + `GeometriaPlantillaEstandarizadaT0Tests` rompen primero (fail-fast `ERR-PLANTILLA`); re-congelamiento por T0 nuevo, nunca edicion silenciosa de literales. Allow-list Phase 0 vigente (68 celdas) intacta. |
| R-JULIO-identidad | Con geometria de julio (delta 0) la recomposicion/llenado debe reproducir byte-identica la formula y los valores canonicos; cualquier deriva es regresion. | Gate julio-identidad: `<f>` byte-identica + goldens fila-a-fila + F680=12.033.011.685,71 + DetRetri-D vs R10. Si julio difiere, STOP (no se avanza a agosto/Q1). |
| R-COLUMNA-fantasma | Escribir una columna que la plantilla no trae o que es formula (Exonera-*, Z sin slot, G..AP con `<f>`). | Guard anti-formula intacto (`CellFormula` -> `ERR-PLANTILLA` nombrando hoja+celda); allow-list de columnas por encabezado resuelto del template; Exonera-* sin contraparte = intactas por construccion. |
| R-MODO-comun | Autocontroles ciegos al modo comun (residual 24.329,38: entra en ambos lados o en ninguno). | Gate externo DetRetri-D vs R10 por ASE explicito (ya existe la costura); no se sella TRUE donde la divergencia es real; tabla observado-vs-esperado por celda en F2. |
| R-FALSO-verde | Apagar/sellar/re-etiquetar un gate para pasar (criterio E1-1..E1-5 Plan 40). | Prohibido `SeGatea=false` como forma de pasar; sub-bloque no evaluable = hallazgo/defecto con frente propio, nunca verde. |
| R-Q1-contagio | Tocar Q1 al arreglar Q2 (fallback posicional prohibido E1-1). | Q1 intacto: 8 rojos declarados siguen rojos; ningun cambio de produccion recibe periodo (D-H); tests Q1 ni se tocan ni se reverdecen. |

---

## 2. DESIGN

### 2.1 Contrato de llenado R1 (clave + columnas)

**Clave de fila (localizador puro, Core, sin I/O, sin periodo — D-H):**

```
clave(fila) = ( C-por-sinonimia , D , E , seccion-por-frontera )
```

- **C-por-sinonimia:** `SinonimosEmpresaR1.SonMismaEmpresa(fuente.C, slot.C)` (Plan 34, catalogo abierto). Cero ramas por nombre de empresa. Normalizacion mayusculas invariantes + sin diacriticos (misma regla del mutador). Roulo no tabulado solo se iguala a si mismo.
- **D (canal):** igualdad exacta `OrdinalIgnoreCase` tras `Trim` (H, O, T, E, 1, 2, 5). Sin sinonimia (no hay catalogo de canales; si el T0 encuentra variante, propone fila citada nueva, no logica ad hoc).
- **E (concepto):** igualdad exacta `OrdinalIgnoreCase` tras `Trim` (Vlr Servicio, Vlr Intereses, Total; mas Mes / Aplicacion nuevos x reversion / Subsidio / Subs-Cont / Componente / AFaseo como roles-frontera). Sin sinonimia.
- **Seccion-por-frontera:** `SeccionEspejoR1.PorFronteraArriba(filas, i)` (Plan 41/T1): `Mes`->Oportuno, `Aplicacion`->Aplicacion, sin frontera arriba->1.a seccion; `Ordinal` distingue ocurrencias (Aplic1/Aplic2, Oportuno1/Oportuno2). `R1FirmaInterior.EsAplicacionPorFrontera` (frontera por DEBAJO, Plan 35-T2) queda INTACTO (otro consumidor, otra semantica). Solo vinculante en filas-dato de empresa (`R1FirmaInterior.EsDatoEmpresa` exige C no vacia); si algun lado es null la seccion no veta (comportamiento historico para Mes/AFaseo/Componente/SubsCont/Total) — salvo lo que F0 decida para filas sin C (ver 2.2).
- **Afinidad de grupo:** el writer agrupa slots por **(firma A/B + clave + seccion)** (`ConstruirGruposSeccion` extendido); cursor INDEPENDIENTE por grupo (doctrina Plan 41: un slot intercalado sigue disponible para una fila emitida despues). `LeerSlots`/`RangoDatosAse` (ASE1 6-72, ASE2 110-210, ASE3 289-355, ASE4 450-550, ASE5 603-676) intactos en geometria; se extiende la lectura de celdas de F/L a **F..Z por encabezado** (ver 2.3).

**Fail-fast:** fila cuya (clave, seccion) no tiene slot en el contrato T0 -> `ERR-SIN-SLOT` nombrando ASE + firma A|B|C|D|E + rotulo de empresa + canal + concepto + seccion esperada + ocurrencia. Prohibido resolver por proximidad, orden fisico o heuristica posicional (E1-1). Fila sin contraparte en fuente -> slot queda en cero predefinido (no se escribe, no es error).

### 2.2 Regla para filas sin C (decision del T0 con numeros, no supuesta)

Opciones que el T0 arbitra con bytes (filas 6-17 ASE1, 619-628 ASE5 + matriz completa por ASE/cuadro):

- **Opcion I (inferencia por grupo vecino):** la fila sin C hereda la empresa del grupo vecino delimitado por fronteras (Mes/Aplicacion/Total) al que pertenece por posicion en la secuencia espejo; si el grupo es ambiguo (cero o 2+ empresas candidatas) -> fail-fast `ERR-SIN-SLOT` con causa `grupo-ambiguo` (nunca fallback silencioso).
- **Opcion F (fail-fast puro):** toda fila dato sin C resoluble por (D,E,seccion) unica -> coloca; si (D,E,seccion) tiene 2+ slots candidatos (uno por empresa) -> `ERR-SIN-SLOT` (nunca por orden fisico).

El T0 decide I vs F con numeros: si la inferencia por grupo vecino reproduce 100% de las colocaciones medidas correctas (F604/F610/F616/F620/F623/F626/F631/F635/F639/F613/F648) sin excepcion y sin introducir ramas por empresa/periodo, se adopta I; en cualquier otro caso se adopta F. **Fallback por orden fisico prohibido en ambas** (E1-1). El micro-punto F654/F661 se aisla en el T0 (emision del lector vs consumo del escritor) y su veredicto condiciona la regla (si el duplicado nace en el lector, la regla no lo tapa).

### 2.3 Colocacion dinamica multicolumna values-only

- **Lectura fuente por encabezado:** el reader R1 resuelve columnas por texto de encabezado (Total, TDF, TTL, TVIAT, Aprovechamiento, CCSA Prest.Aprov., SERVICIO ESPECIALES, TCS, TLU, TBL, TRT, CCSA No Aprov., Ajuste Decena, Ajuste Recargo, Intereses Cte, Rec.Intereses), no por letra fija. Fuente trae F..U (T=Intereses Cte, U=Rec.Intereses).
- **Escritura plantilla por encabezado:** el writer resuelve columnas destino por texto de encabezado del template (F..Z). Mapeo: cada encabezado fuente -> mismo encabezado plantilla; **Rec.Intereses U(fuente) -> Z(plantilla)**; Exonera-* (U/V/W/X plantilla) sin contraparte = intactas; T/Y plantilla (Intereses Cte, doble columna) reciben el mismo valor fuente T (el T0 confirma si T y Y son espejo o si Y tiene semantica propia; si Y es espejo, ambas reciben T; si no, el T0 lo declara y el writer solo toca la columna con encabezado identico).
- **Values-only:** solo `<v>` numerico con decimal exacto de la fuente (sin redondeo; el redondeo a entero es exclusivo de DetRetri, Plan 29 Unidad D). Guard anti-formula intacto: celda destino con `<f>` -> `ERR-PLANTILLA` (nunca sobrescritura). Cero `InsertRow`/`DeleteRow`, cero `new CellFormula`, cero Reanclar en este path (grep de cierre lo demuestra). Slots no consumidos quedan en cero.
- **Rango de slots:** `RangoDatosAse` intacto; F682/F702 (EXTEMPORANEO ASE5, literales 0) siguen FUERA del rango salvo arbitrio UAESP (residual 24.329,38: ver F2; sin arbitrio no se incluyen).

### 2.4 Que filas/columnas NO se tocan

- Anclas y visibles (`F75/F254/F415/F568/F680`, `F80..F705`, `F700`, `F682/F702`, `CONSOLIDADO!D9:D13/D51/D108/D109`, `REMUNERACION_*!D9:D13/D51/D108`, `VALIDACION_*`, `DetRetri/DetValiRetri`, `INTERVENTORIA`, `N3/G7/K7/D6/D7`, Phase 0 allow-list de 68 celdas): **no se tocan** en F1 (las recompone/evalua su propio dueno: Plan 31 recomposicion, Plan 28 sello, Plan 39/40 Phase 0).
- Columnas Exonera-* (U/V/W/X) sin contraparte: intactas.
- Fuera de los bloques espejo R1 (`RangoDatosAse`): prohibicion insert/delete intacta (derogacion Plan 21 solo para bloques espejo, y este path ni siquiera la usa: es values-only por slot).
- Q1: no se toca (ni geometria, ni writer, ni tests).

---

## 3. SPEC (criterios por tarea + gates)

### S0 - Gates transversales (toda T)

- G-WORKBOOK: todo gate de negocio lee el **workbook generado** (texto-`<f>` + literales; prohibido `<v>` stale como oraculo, doctrina Plan 28/29) y lo contrasta contra el dominio por firma ±0.5.
- G-F-BYTE: `<f>` byte-identica en julio-identidad (recomposicion Plan 31 + shared-formula Plan 33: `EscribirFormulaPreservando`, invariante `si` workbook-wide).
- G-DETRITRI-R10: DetRetri calculado bottom-up (`ROUND(D104:D108,0)`) vs R10 (`DetRetri{AAAAMMQ}` D9:D13 + D14) ±0.5 post-redondeo en AMBAS quincenas; divergencia -> fail-fast con periodo + archivo + ambos valores. Intacto.
- G-AUDITORIA (al llegar): la auditoria manual de la propietaria es el **oraculo de paridad** (geometria + ejemplos congelados por F0); sus veredictos entran como filas citadas (nuevas divergencias declaradas con 6 campos o defects con frente propio). Sin auditoria no hay verde por ausencia.
- G-SHA: oficial SHA256 intacto antes/despues (`3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).
- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` 0 warnings / 0 errores; CRLF sin mezclar; NUNCA `git commit`/`push`; rutas explicitas; regens a temp FRESCA con el OFICIAL como unica entrada; prohibido `Docs/Insumos` raiz en lo nuevo; tests NO se debilitan (0 `Skip`, 0 `Assert.True(true)`, 0 `NotImplementedException`).

### S1 - Criterios F0 (T0 evidencia read-only)

- S1.1: fixture read-only cita bytes del oficial (rangos por ASE, encabezados F..Z, filas 6-17 / 110-121 / 289-302 / 450-461 / 603-616 sin `Vlr Intereses`, slots 612/628/647/659/673 + seccion por frontera, `F700` 3 terminos, `F682/F702` literales 0, `F680` 5 terminos) + del manual legacy (O 402,40 / T 7.697,23 en ASE5; ASE1 sin Intereses de 1.a) + de la fuente julio-2 (filas 8/11 O/T con desglose, filas OCCIDENTE 13/32/39/53, filas sin C medidas).
- S1.2: matriz (C,D,E,seccion)->slot por ASE/cuadro publicada en el doc T0; cada celda con cita (fila fuente + celda slot + seccion).
- S1.3: veredicto I vs F (2.2) con numeros (reproduccion 11/11 de las colocaciones medidas o fail-fast justificado); veredicto T/Y (espejo o semantica propia); aislamiento F654/F661 (lector vs escritor) con cita.
- S1.4: puerta auditoria: geometria + ejemplos congelados + prueba de paridad redactada pero PENDIENTE (no ejecutada sin la entrega).

### S2 - Criterios F1 (implementacion)

- S2.1: fila fuente con desglose -> slot con **todos sus componentes** por encabezado (F + G..Z segun mapa; U->Z; Exonera-* intactas). Test fila-a-fila: OCCIDENTE 1.a (Total 417.519,62, TDF 26.657,43, TTL 14.138,02, TVIAT 7.723,27, Aprov 93.911,73, CCSA 2.290,29, TCS 20.089,34, TLU 28.092,36, ...) en los 5 ASE (muestra + barrido).
- S2.2: filas sin C colocadas por la regla F0 (I o F) sin orden fisico; test de regresion de las 11 colocaciones medidas ASE5 (F604/F610/F616/F620/F623/F626/F631/F635/F639/F613/F648 + 674 en 0).
- S2.3: ausencia de fuente = cero predefinido intacto (Vlr Intereses H en 0, AFaseo Area Limpia vacio) sin `ERR-SIN-SLOT`; fuente sin slot = `ERR-SIN-SLOT` con ASE+firma+empresa+canal+concepto+seccion+ocurrencia.
- S2.4: `<f>` intacta en todo el diff F1 (grep: cero `new CellFormula` fuera de `EscribirFormulaPreservando`; cero `InsertRow`/`DeleteRow`); D-H (cero literales/ramas por periodo o empresa en produccion); tolerancia ±0.5 solo en comparacion, nunca en escritura (decimal exacto).

### S3 - Criterios F2 (re-verificacion autocontroles/divergencia)

- S3.1: tabla fresca observado-vs-esperado-Δ por celda (O25/D34/F34/C15/D25) + DetRetri-D vs R10 por ASE, evaluada desde lo escrito (nunca `<v>`).
- S3.2: re-decision Via A/B con evidencia: A = mantener OCCIDENTE-ASE5 (Δ 417.519,62, dueno UAESP) y residual como hallazgo para arbitrio; B = ampliar/retirar por condicion de retiro solo con cita UAESP/auditoria. Sin evidencia fresca no hay cambio de via.
- S3.3: declarada = ultima puerta (D-F Plan 42): si hay CUALQUIER no declarada, FAIL duro aunque la declarada este reconocida; doble huella obligatoria (reporte durable + log) si hay reconocimiento; codigo de salida 6 distinto de 0.

### S4 - Criterios F3 (corridas frescas Julio-2)

- S4.1: corrida Q2 (2026072) desde el OFICIAL a temp FRESCA (rutas explicitas; plantilla CERRADA en Excel); F680=12.033.011.685,71; D34/F34/O25 a 417.519,62 exactos salvo veredicto F2; C15/D25 TRUE; DetRetri-D vs R10 intacto.
- S4.2: suite completa con trx + set-diff vs linea base **39 Failed / 873 Passed / 912 Total** (mismos 39; 8 de Q1 declarados siguen rojos por su causa); 0 warnings; SHA256 intacto.

---

## 4. TASKS

### T0 - Evidencia read-only con fixtures (F0; cero codigo de produccion; bloqueante)

1. T0.1 Inventario BCL (ZIP+XML, sin Excel/COM/OpenXML): encabezados F..Z del oficial por bloque ASE + encabezados F..U de la fuente R1 ASE5 (y muestra ASE1) + filas 6-17 (ASE1) y 619-628 (ASE5) + filas fuente 8/11 (O 402,40 / T 7.697,23 con desglose) + filas sin C medidas + `F654`/`F661` + `F612`/`F613`/`F680`/`F700`/`F682`/`F702` + `F628` huerfano (grep de consumidores en todas las hojas, maestros shared incluidos). Rutas explicitas: `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx`, `Docs/Prueba Julio-2/Insumos/5-Area Limpia/Recaudoporcomponente_..._20268465657830.xlsx`, `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx`, `Docs/Prueba Julio-2/Resultado/6/Remuneración 202607-2 Total.xlsx`. Prohibido `Docs/Insumos` raiz. SHA256 antes/despues.
2. T0.2 Matriz (C,D,E,seccion)->slot por ASE/cuadro + veredicto I vs F (2.2) + veredicto T/Y + aislamiento F654/F661 (lector vs escritor). Fixture: `Remuneracion.IntegrationTests/R1FuenteCanalConceptoT0Tests.cs` (solo lectura + asserts de bytes; no toca produccion).
3. T0.3 Congelar geometria + ejemplos del Ingeniero y redactar la prueba de paridad contra la auditoria (PENDIENTE de su entrega; puerta explicita). Mapeo por ASE/cuadro del hueco `Vlr Intereses` (ASE5 con intereses en la anterior; ASE1 sin ellos) + reporte al dueno (texto listo, sin enviar plantilla).
4. Gate T0: fixture verde + doc `Docs/T0-Plan44-R1-Fuente-Canal-Concepto.md` con tablas citadas + decisiones I/F, T/Y, F654/F661 + puerta auditoria. Sin esto no se abre T1.

### T1 - Implementacion (F1; contract-first)

1. T1.1 Core puro: extender la clave a `(C-por-sinonimia, D, E, seccion)` (predicados puros, sin I/O, sin periodo; catalogo abierto; D-H). Incluye regla sin-C segun veredicto T0 (I o F) + `ERR-SIN-SLOT` enriquecido (ASE+firma+empresa+canal+concepto+seccion+ocurrencia).
2. T1.2 Reader R1 por encabezado (F..U fuente; T=Intereses Cte, U=Rec.Intereses) + writer slot multicolumna values-only (F..Z por encabezado; U->Z; Exonera-* intactas; T/Y segun veredicto T0; decimal exacto; guard anti-formula; cero insert/delete). `RangoDatosAse` y `LeerSlots` (geometria) intactos salvo extension de columnas.
3. T1.3 Cableado (ProcesadorPeriodo/ProcesadorRemuneracion + mapa leaf R1-Q2 por firma sobre la secuencia espejo, Plan 25 intacto): el agregado `totOpt`/`extemp` (dominio por firma) sigue mandando en CONSOLIDADO; el interior multicolumna alimenta las anclas por empresa. Sin ramas por periodo/empresa.
4. Gate T1: tests focalizados fila-a-fila (S2.1/S2.2/S2.3) verdes + build 0 warnings + greps (0 `<f>` fuera de Phase 0/`EscribirFormulaPreservando`, 0 Insert/Delete, 0 literales por periodo/empresa).

### T2 - Revision + 0 warnings + set-diff (calidad senior)

1. T2.1 Code-review independiente (spec compliance + calidad): trazabilidad cada cambio -> requisito S2; verificacion de que ningun test fue debilitado; CRLF sin mezclar; 0 recursos UI inventados (no hay UI en este plan).
2. T2.2 Suite completa con trx + set-diff vs linea base **39 Failed / 873 Passed / 912 Total** (cero rojos nuevos; 8 de Q1 declarados intactos). Comando: `dotnet test Remuneracion.WinForms/Remuneracion.WinForms.slnx --logger "trx" ` con `--filter` focalizados primero.
3. Gate T2: review APROBADO (o APROBADO CON OBSERVACIONES cerradas) + set-diff vacio + SHA256 intacto. Sin esto no se abre T3.

### T3 - Corridas frescas Julio-2 desde el OFICIAL a temp FRESCA + re-verificacion (F2+F3)

1. T3.1 Regen Q2 (2026072): `dotnet run --project Remuneracion.Cli -- --periodo 2026072 --carpeta <insumos-julio-2> --plantilla <oficial> --salida <temp-fresca>` (rutas explicitas; plantilla CERRADA). Repetir sin/con `--aceptar-divergencia-declarada OCCIDENTE-ASE5` segun veredicto F2 (exit 1 vs exit 6 + reporte + log).
2. T3.2 Re-verificacion F2 (S3.1/S3.2/S3.3): tabla fresca O25/D34/F34/C15/D25 + DetRetri-D vs R10 por ASE + re-decision Via A/B con evidencia (sin presuponerla). Si la auditoria ya llego, ejecutar la prueba de paridad F0 (oraculo); si no, dejarla PENDIENTE con constancia.
3. Gate T3: S4.1/S4.2 verdes + reporte durable (si aplica) + SHA256 intacto. Cualquier autocontrol movido por el re-llenado queda explicado celda-a-celda (no se apaga ni se sella).

### T4 - Docs (project-context.md + manual)

1. T4.1 Actualizar `.opencode/project-context.md` (doctrina Plan 44: contrato de llenado + regla sin-C + mapa U->Z + veredictos T0 + via A/B vigente + puerta auditoria).
2. T4.2 Nota al manual de HU-17 (instructivo Capa B): que auditar fila-a-fila por (fuente, canal, concepto, seccion) y donde mirar el reporte de divergencias declaradas.
3. Gate T4: diff de docs citado en el resumen de cierre; cero codigo de produccion en T4.

---

## 5. Invariantes vigentes (todos heredados, sin excepcion)

Values-only; jamas `<f>` fuera de Phase 0 (allow-list 68 celdas: 16 Plan 39 + 36 recaudo + 16 F-NEW-1); cero `InsertRow`/`DeleteRow`; preservacion shared-formula (Plan 33); D-H (cero literales/ramas por periodo o empresa); tolerancia ±0.5 solo en comparacion; CRLF sin mezclar; NUNCA `git commit`/`push`; oficial SHA256 intacto antes/despues; quincena por dominio (`Periodo.NumeroQuincena`), nunca por contenido; tests NO se debilitan; declarada = ultima puerta con doble huella y fail-closed (Plan 42); prohibido fallback posicional (Plan 41 E1-1) y prohibido sellar/apagar/re-etiquetar (Plan 40 E1-1..E1-5); Q1 (8 rojos) ni se toca ni se reverdece; `<v>` nunca es oraculo.

Comandos:

```bash
dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx
dotnet test Remuneracion.WinForms/Remuneracion.WinForms.slnx --filter "<focalizados>" --logger "trx"
dotnet test Remuneracion.WinForms/Remuneracion.WinForms.slnx --logger "trx"
```

Rutas explicitas en todo lo nuevo; prohibido `Docs/Insumos` raiz; regens siempre a temp FRESCA con el OFICIAL como unica entrada y la plantilla CERRADA en Excel.

---

## 6. Architecture Validation Certificate

- **SOLID:** SRP (clave pura en Core / lectura por encabezado en Infrastructure-reader / pegado values-only en writer / clasificacion en `ValidadorBasico` + catalogo de datos); OCP (nueva empresa/canal/concepto/columna = fila citada en catalogo o encabezado resuelto, cero ramas nuevas; periodo nuevo = cero cambios por D-H); LSP/DIP (writer depende de `BloqueEspejoAseInputs` + predicados, no de concretos de periodo); ISP (contratos `IRecaudoReader`/`IPlantillaWriter`/`IDetRetriR10Reader`/`IReporteDivergenciasDeclaradas` intactos).
- **Best Practices:** contract-first (T0 congela bytes antes de T1); fail-fast accionable (`ERR-SIN-SLOT`/`ERR-PLANTILLA` con ASE+firma+seccion+celda); gates que leen el workbook (texto-`<f>` + literales, nunca `<v>`); goldens ±0.5 + `<f>` byte-identica + DetRetri-D vs R10; set-diff contra linea base; SHA256 del oficial; CRLF; sin commits.
- **Performance:** resolucion por encabezado una vez por hoja (mapa encabezado->indice cacheado por corrida); agrupacion por (firma,seccion) con cursor por grupo (O(filas + slots)); lectura acotada en T0 (BCL ZIP+XML); sin abrir workbooks en preflight; regens a temp fresca sin copiar de mas.

---

## 7. Clarification Gate (que necesita aprobacion del Ingeniero)

1. **APROBADO 2026-10-16** (documento + ejecución F0→F3 con Enmienda E-44.1).
2. **RATIFICADO P2 + cero documentado 2026-10-16** (Enmienda E-44.1): I vía contrato explícito por ASE; fallback prohibido.
3. **Via A/B tras T3** con evidencia fresca (mantener OCCIDENTE-ASE5 / ampliar / retirar solo por condicion de retiro UAESP).
4. **Puerta auditoria:** cuando llegue la auditoria de la propietaria, habilitar la prueba de paridad F0 como oraculo (nuevo T0 de reconciliacion si discrepa de lo congelado).
5. **Reporte del hueco `Vlr Intereses`** al dueno de la plantilla (texto del T0.3; NO tocar la plantilla).

Sin estas aprobaciones el plan no avanza de T0 a T1 (1-2) ni cierra F2 (3-5).
