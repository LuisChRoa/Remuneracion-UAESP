---
titulo: Plan 59 - Eliminacion de coordenadas quemadas con geometria derivada de la plantilla v3
estado: APROBADO 2026-10-26 — Clarification Gate §6 ratificado 8/8 por el Ingeniero ("Si procede con calidad senior"); H0 habilitado
fecha: 2026-10-26
numero: 59
planes_ocupados: 01..58 ocupados; este es el 59
padre: plans/58 - integridad-paquete-y-conciliacion-r2-por-empresa.md
madre_doctrina: .opencode/project-context.md (plantilla autoridad; values-only; guard anti-formula; tolerancia ±0,5; prohibido <v> stale como oraculo) + plans/55 (conciliacion por empresa, anti-tautologia, candado) + plans/56 (identidad A/B, cero documentado, candado de dos capas) + plans/57 (anclas derivadas, precedente de eliminacion) + plans/40 (W-1 evalua+sella) + plans/31 (gate workbook-vs-dominio)
caso_rojo: corrida real 2026072 --cinco-ase contra la v3 — ERROR [ERR-SIN-SLOT]: ASE 1: la fila A|B|C|D|E='Subs/Cont|Total|||' de la fuente R1 no tiene destino en la seccion 'Subs/Cont' (hoja 'Reporte Componentes R1', rango del bloque 6-72); cascada ASE2..ASE5 en VerificarCabeceraBloque (B108/B287/B446/B599 vacias en v3). Una sola causa, no cinco.
gate_entrada:
  plantilla_v3: Docs/Plantilla estandarizada/v3/Formato Remuneracion 2026xx-x Total v3.xlsx — SHA256 8199F8647A7ACA84812AD3A61F6607956B2270E0B5C1C999C7690B62032B1295 (705.934 bytes; verificar intacto antes/despues de cada H)
  plantilla_v2: historica desde 2026-10-26 (SHA 0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB intacto, no se consume)
  plantilla_v1: historica (SHA 3086AC31...C5214 intacto, no se consume)
  suite_base: ratificada 2026-10-26 = corrida del arbol VIGENTE sin el fixture de H0 (1558 total / 1518 verdes / 40 rojos / 0 skip), trx `Remuneracion.IntegrationTests/TestResults/p59-h0-baseline-hoy.trx`. Los 6 verdes->rojo y 1 rojo->verde frente a `p58-final.trx` tienen causa externa citada: commit `88b3acb` "Actualizacion de insumos Julio-2" (2026-10-26 07:46) reemplazo 46 archivos de `Docs/Prueba Julio-2/Insumos` por versiones corregidas (los insumos anteriores "no eran del todo correctos"); el Ingeniero confirma que los VIGENTES son los de prueba. Rojo nuevo fuera de esta lista = STOP
  h0_estado: CERRADO 2026-10-26 - 5/5 frentes PASS, 0 discrepancias vs §0.4; fixture `GeometriaDerivadaV3T0Plan59Tests` 23/23 verde; build 0 warnings; evidencia `Docs/T0-Plan59-Geometria-Derivada-V3.md`
  d2_ratificado: el inicio de bloque se fija por la CABECERA DE BLOQUE (columna B = rotulo del ASE del catalogo de dominio), no por "fila siguiente al gran total anterior" (ese hueco contiene el cuadro-resumen con firmas de cierre reutilizadas)
  sin_commits: NUNCA git commit / git push; NUNCA git checkout / git reset
  excel_cerrado: Excel debe estar cerrado durante las corridas
---

# Plan 59 — Eliminación de coordenadas quemadas con geometría derivada de la plantilla v3

> SDD completo PROPOSE → DESIGN → SPEC → TASKS. Fuente única de verdad para la implementación una vez aprobado ESTE documento. No se implementa sin aprobación del Ingeniero (punto de bloqueo: sin su OK no se abre H0; estado inicial PROPUESTO).
> Lectura obligatoria previa: `plans/58` (padre: P1 resuelto por la v3, P2 abierto) + `plans/56` (identidad A/B, §7/§8) + `plans/55` (conciliación por empresa, candado) + `plans/57` (precedente de eliminación) + `.opencode/project-context.md` (doctrina madre).
> Entorno: Windows, repo con espacios en la ruta; plantillas CERRADAS en Excel al correr tests (FileShare). Stack .NET 10 (Core/Infrastructure/WinForms) + Excel; persistencia por archivos Excel, NO base de datos. Español; CRLF sin mezclar.
> Visual Design Intent: N/A (cero UI en este plan).

---

## 0. Mandato vinculante, evidencia medida y alcance

### 0.1 Norma suprema del plan (textual del Ingeniero, citada íntegra)

> "No debemos tener filas quemadas en la App, en su defecto manejarlo bajo los conceptos que ya sabemos de la plantilla, principalmente las empresas (ASE1, ASE2, ASE3, ASE4 y ASE5) y como lo son, secciones, subsecciones, fuentes de recaudo, canales de recaudo, conceptos y componentes."

### 0.2 Consecuencias obligatorias (el plan las hace cumplir con tests de candado)

1. **PROHIBIDO** cualquier literal de fila/celda/ventana de fila en producción (`6`, `72`, `F682`, `E285`, `"F53"`, rangos `(6,72)`, offsets aritméticos, mapas congelados por (empresa,ASE)). El único mecanismo de localización admitido es el cruce de **conceptos**: ASE → sección (A) → subsección (B) → ocurrencia de B dentro de A → fuente de recaudo (C, con catálogo abierto de sinónimos) → canal (D) → concepto (E) → componente por encabezado de columna.
2. **PROHIBIDO** ramificar por empresa, ASE, período o quincena (`if (ase==2)`, `if (hoja=="INTERVENTORIA")`, `if (periodo.Codigo=="2026072")`, y toda forma equivalente escondida en mapas, allow-lists o mensajes).
3. **PROHIBIDO** insertar/eliminar filas o columnas y **PROHIBIDO** escribir `<f>`: solo valores (`<v>`) en celdas de valor del template. El guard anti-fórmula (`GuardEscrituraFormula`, `EscribirValorNumerico`) se mantiene intacto.
4. Un template que no declare el destino exigido ⇒ **fail-fast accionable** (`ERR-SIN-SLOT` / `ERR-PLANTILLA`) nombrando bloque+firma+sección; nunca reacomodo por posición ni remapeo silencioso.
5. Calidad senior: la solución debe **eliminar la clase de defecto** (coordenadas quemadas), no re-congelar una foto nueva de la v3.

### 0.3 La v3 es la ÚNICA plantilla vigente desde 2026-10-26

Archivo: `Docs/Plantilla estandarizada/v3/Formato Remuneracion 2026xx-x Total v3.xlsx` — SHA256 `8199F8647A7ACA84812AD3A61F6607956B2270E0B5C1C999C7690B62032B1295` (705.934 bytes). v2 queda histórica (SHA `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB`), v1 histórica (`3086AC31...C5214`). La SHA de la v3 se re-verifica intacta en el T0 y antes/después de cada H.

### 0.4 Evidencia medida por bytes (verificada por el orquestador con dump OpenXML propio; el plan la cita como "medida" y el T0 la re-verifica)

**a) v3 es superconjunto estricto de v2.** Comparación multiconjunto de etiquetas (A,B,C,D,E) por bloque y subsección en `Reporte Componentes R1`, `Rem. Anticipos R2` y `Reversion Pagos R4`: **0 faltantes** en las tres hojas. Mismas 40 hojas, mismos nombres y orden. Mismas columnas (R1 `A1:AP`, R2 `A1:S`, R4 `A1:R`). Cero columnas nuevas o eliminadas.

**b) Geometría (filas de cabecera de bloque por ASE):**

| Hoja | v2 (total filas / cabeceras) | v3 (total filas / cabeceras) | Δ |
|---|---|---|---|
| `Reporte Componentes R1` | 712 / 4,108,287,448,601 | **843** / 4,**122**,**356**,**517**,**725** | +131 (ASE1 +14, ASE2 +55, ASE3 0, ASE4 +55, ASE5 +7) |
| `Rem. Anticipos R2` | 485 / 1,65,173,291,402 | **527** / 1,**75**,**196**,**317**,**441** | +42 |
| `Reversion Pagos R4` | 382 / 1,91,192,229,344 | **394** / 1,**97**,**198**,**235**,**350** | +12 |

**c) Novedades de negocio en v3** (todas verificadas contra las fuentes reales):
- `NUEVO ESQUEMA` (EAAB Reciprocidad) pasa de canales `{1,2,5}` (R1) / `{1,2,3,5,9}` (R2) a **`{1,2,3,4,5,6,9}`** en ASE2 y ASE4, en los 4 cuadros de R1 (AFaseo/Componente × Mes/Aplicacion nuevos x reversion) y en los 2 de R2.
- `ENERBIT` gana slot de canal `E` en R1 y R2 de **ASE1** (en v2 ASE1 solo tenía ENEL(E) + OCCIDENTE(H,O,T)); ASE5 se normaliza (en v2 el par `E/ENERBIT` solo existía en el cuadro "Aplicacion").
- `Reversion Pagos R4` gana bloque por empresa `A=ENERBIT` en **ASE1** y **ASE5** (+6 filas cada uno). R4 no tiene dimensión de canal.
- `Rem. Anticipos R2`: el 2.º cuadro gana el grupo `canal A / empresa EAB` en los **5 ASE** (en v2 solo lo tenía el 1.er cuadro).
- Las zonas nuevas tienen **0 celdas con `<f>`**: son slots de valor puros ⇒ la regla "solo datos, nunca fórmulas" es satisfacible sin tocar la plantilla.
- `CONSOLIDADO_TOTAL RECAUDO` conserva geometría y celdas de escritura (D9:D13, D28:D32, D47:D51, D66:D70, D85:D89, D104:D109). 1.993 fórmulas fueron re-apuntadas por Excel a las filas nuevas con **forma idéntica** (CONSOLIDADO, DetRetri/DetValiRetri, Informe AFaseo, 5 × `REMUNERACION_*`). Sin cambios: `Recaudo *`, `REPORTE RECAUDO x BANCO`, `ANT*`, `BCE SC POR FACT.`, `SALDOS POR NOTA`, `RETRIBUCION NEGATIVA`, `AJUSTES - SF-T`, `INTERVENTORIA`, `VALIDACION_*`, `GERENTES_*`.

**d) Semántica de etiqueta confirmada (clave del diseño):** en R1 y R2 la fila `C=<empresa> / D=Total` (R1) o `B=<empresa> / C=Total` (R2) es **CIERRE (trailer)** del grupo de canales que la **precede**. Verificado con fuente real: ASE2 octubre‑1 `Recaudoporcomponente`, donde el total `OCCIDENTE` = `H+O+T` al centavo. R4 no compone: su rótulo de grupo (columna A) ya es la empresa.

**e) Canales que los insumos realmente traen** (20 fuentes `Recaudoporcomponente`: Julio‑1, Julio‑2, Agosto‑2, Octubre‑1 × 5 ASE): los canales numéricos solo existen bajo `NUEVO ESQUEMA` y son `{1,5}` (Q1 y octubre‑1) o `{1,2,5}` (Q2). **Nunca aparecen 3, 4, 6 ni 9.** Los slots nuevos son **capacidad de plantilla, no demanda**: la fuente que no los trae cae en el 0 del template con huella (doctrina Plan 50). Igual el nuevo slot `E/ENERBIT` de ASE1 (ASE1 solo factura ENEL+OCCIDENTE en los 4 períodos medidos).

**f) El defecto que motiva el plan (corrida real 2026072 --cinco-ase con la v3):**

```
ERROR [ERR-SIN-SLOT]: ASE 1: la fila A|B|C|D|E='Subs/Cont|Total|||' de la fuente R1
no tiene destino en la sección 'Subs/Cont' ... hoja 'Reporte Componentes R1', rango del bloque 6-72.
```

Causa raíz medida: la fila de cierre `A=Subs/Cont / B=Total` del bloque ASE1 está en `R1!A71/B71` en la v2 (**dentro** de la ventana congelada 6‑72) y en `R1!A85/B85` en la v3 (**fuera**). La app solo mira 67 filas de un bloque de 118 (ASE1 va de la fila 4 a la 121). La plantilla **sí** declara el destino; el error es 100% de la app. Cascada: ASE1 falla por ventana corta; ASE2..ASE5 fallarían antes en `VerificarCabeceraBloque` porque `B108/B287/B446/B599` quedaron vacías en la v3. **Una sola causa, no cinco.**

**g) Coordenadas quemadas inventariadas en producción (el alcance real a eliminar):**

| Artefacto | Archivo:línea | Contenido | Efecto con v3 |
|---|---|---|---|
| `RangoDatosAse` | `Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs:82` (+ uso en `:126`, `VerificarCabeceraBloque`) | 5 ventanas `{1:(6,72) 2:(110,210) 3:(289,355) 4:(450,550) 5:(603,676)}` | `ERR-SIN-SLOT` ASE1; `ERR-PLANTILLA` ASE2..5 |
| `WorkbookLeafCellMapQ2` | `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapQ2.cs` — `R1Q2EditablesPorAse:75`, `R1Q2ProtectedPorAse:132`, `R2Q2EditablesPorAse:169`, `R2Q2ProtectedPorAse:182`, `R4Q2EditablesPorAse:196`, `R4Q2ProtectedPorAse:209`, `DetRetriDestinoPorAse:264` | 49 celdas literales (p.ej. `F53`, `F55`, `F206`, `F560`; cadenas de fórmula literales `F75=F47+F70+F16-L16-L47`, `F77`, `E47`, `D73`) | `ERR-PLANTILLA` o escritura en fila semánticamente distinta |
| `ProtectedFormulasPorEmpresa` | `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapPorEmpresa.cs`, consumido en `OpenXmlPlantillaWriter.cs:816` | 78 celdas literales por (empresa,ASE) | `ERR-PLANTILLA` |
| Guard de validación | `OpenXmlPlantillaWriter.cs:570` (`ValidarCeldaTieneFormula`) + llamadas literales en `:83-84, :102, :119, :544-547` | compara celda + fragmentos literales | rechazo de una plantilla legítima |

Ya está resuelto y **sirve de precedente de diseño** (no reabrir): `AnclasExtempR1` (Plan 49) fue eliminado por el Plan 57 mediante derivación de anclas desde la plantilla; `AnclasExtempR1.cs` ya no existe. El camino label‑driven que **sí** es tolerante y debe reutilizarse: `ContratoSeccionPeriodFreeR1` (`PlanificarJerarquiaR1`, `SeccionPorCierre`, `FirmaCuadroR1`, `CanalEfectivo`, `Alinear`), `SinonimosEmpresaR1`, `ResolvedorConciliacionEmpresa` (Plan 55), `EscritorValoresSlotR1` (identidad A/B/C/D/E del Plan 56), `ResolvedorOperandosEmpresaR1` (Plan 53).

### 0.5 Hallazgo que cambia alcance (la v3 resuelve el pendiente P1 del Plan 58)

La v3 **resuelve el pendiente P1 del Plan 58** (`EAB` −18.613,87 en ASE2, `VALIDACION_TOTAL!D34` = FALSE). Evidencia: en `RerpoteDetalleSaldosaFavor` de ASE2 (octubre‑1) hay **dos** grupos `C=A … B=EAB`: `+149.137,87` (1.er cuadro) y `−18.613,87` (2.º cuadro). El 2.º cuadro de R2 en la v2 **no declaraba** el grupo `A/EAB`, por eso el valor no tenía destino. La v3 lo declara en los 5 ASE. Además, `R-MAPA-V2-R2-ASE4` (recorte de la conciliación por empresa en Q2, hoy bloqueado porque `R2!E285` es fórmula en la v2) **pierde su bloqueo**: en la v3 esa coordenada ya no es fórmula y el 2.º cuadro declara el grupo. Nota: solo ASE2 tiene grupos `A/EAB` en ese período (0 en ASE1/3/4/5). Queda **abierto y fuera de alcance**: Plan 58‑P2 (columna E de disposición final vacía en R2/R4) — requiere decisión de negocio; este plan lo declara explícitamente como NO tocado.

---

## 1. PROPOSE

### 1.1 Intent

Que la app **deje de conocer filas, celdas y ventanas**: todo destino de valor y toda protección se deriva de la **estructura de etiquetas que la propia plantilla declara** (ASE → sección A → subsección B → ocurrencia → fuente C → canal D → concepto E → componente por encabezado), de modo que la v3 —y cualquier geometría futura con la misma gramática— se procese sin cambiar código. Al cerrar, la conciliación por empresa en Q2 queda **habilitada** (el recorte pierde su bloqueo) y el `EAB` de ASE2 cierra al centavo por construcción.

### 1.2 In Scope (H0 + H1 + H2 + H3, prohibido ampliar sin nuevo OK del Ingeniero)

| ID | Entregable | Criterio de salida |
|---|---|---|
| **H0/T0** | Inventario read-only por bytes, cero código de producción: (a) re-verificación de §0.4 (superconjunto, cabeceras, SHA v3); (b) vocabulario exacto de cierres por hoja (qué fila es cierre de grupo/subsección/sección/bloque en R1/R2/R4); (c) estabilidad de `D9:D14`/`D104:D109` (¿derivables por etiqueta o ancla fija documentada por etiqueta?); (d) mapa etiqueta→celda del 2.º cuadro R2 con `A/EAB` en los 5 ASE | Fixture verde + doc de evidencia; puerta H0→H1 |
| **H1** | Ventanas y cabecera de bloque derivadas (elimina `RangoDatosAse`) | CA-B1 en corrida real |
| **H2** | Destinos de valor y protecciones derivados (elimina mapas congelados de celdas y guard literal) | CA-B2 en corrida real |
| **H3** | Conciliación por empresa en Q2 habilitada (retira el recorte) + verificación `EAB` ASE2 al centavo + re-verificación `R-MAPA-V2-R2-ASE4` | CA-B3 en corrida real |
| **Candado** | Extensión de `ConciliacionEmpresaCandadoAntiQuemadoTests`: falla si aparece en producción cualquier literal de fila/celda/ventana o rama por empresa/ASE/período (incluida la forma escondida en mapas y allow-lists) | CA-B4 |
| **Gate** | Corridas CLI julio‑2 (`2026072 --cinco-ase`), octubre‑1 (`2026101`), julio‑1 (`2026071` no-regresión) + suite completa con set-diff de rojos | CA-B5 |

### 1.3 Non-goals / out (vinculante)

- No tocar la plantilla (SHA v3 intacto antes/después de cada H; cero `<f>`, cero `InsertRow`/`DeleteRow`, cero geometría).
- No tocar fórmulas del workbook generado salvo lo que el diseño autorice con cita (por defecto: nada — values-only).
- Plan 58‑P2 (columna E de disposición final vacía en R2/R4): NO tocado, requiere decisión de negocio.
- No reabrir `AnclasExtempR1` (eliminado por Plan 57), ni la semántica de agregados (Plan 25 D-C/D-F), ni la lectura única R1 (Plan 54), ni el resolvedor por etiquetas (Plan 55), ni la jerarquía A/B (Plan 56).
- No re-congelar tablas por celda/ASE/empresa/quincena/período bajo ninguna forma (ni mapa, ni allow-list de coordenadas, ni "solo lectura").
- Commits, UI, CLI nueva, docs de usuario. NUNCA `git commit` / `git push`. NUNCA `git checkout` / `git reset`.

### 1.4 Riesgos con mitigación

| ID | Riesgo | Mitigación |
|---|---|---|
| R-REFOTO | Re-congelar la v3 (cambiar `(6,72)` por `(4,121)` y listo) — viola §0.2.5 | §2-D0 (diseño que lo prohíbe) + CA-B4 (candado que lo detecta) + alternativa M-B descartada con motivo |
| R-RAMA-OCULTA | La rama por ASE/empresa sobrevive escondida en un mapa o mensaje | Candado §2-D5 inspecciona mapas y allow-lists, no solo `if`; prueba de injerto con mapa escondido |
| R-CIERRE-AMBIGUO | El vocabulario de cierres varía por hoja (R1≠R2≠R4) y el delimitador único no lo cubre | H0 fija el vocabulario por hoja con bytes; el delimitador parametriza por hoja, nunca por ASE |
| R-DETRI-FIJO | `DetRetriDestinoPorAse`/`D104:D109` resultan ser anclas estables y alguien propone "dejarlas" | T0 verifica; si son estables se documentan como derivadas por etiqueta (la etiqueta existe aunque la fila no se mueva) |
| R-ORACULO-V | Usar cachés `<v>` stale de celdas con `<f>` como oráculo | Método H0 y gates: texto-`<f>` + literales + fuente leída; `<v>` de celda con `<f>` nunca es oráculo |
| R-Q2-RECORTE | Al retirar el recorte Q2 reaparece un rojo latente de otra familia | H3 corre con set-diff contra la base; rojo fuera de lista = STOP, el recorte se revierte por config, no con parche |
| R-PLANTILLA | El fix toca una `<f>` o la geometría | SHAs + G-FORMULA + guard anti-fórmula intacto; H0 read-only |

---

## 2. DESIGN

### D-0. Método H0/T0 por bytes (read-only bloqueante, cero código de producción)

**Forma normativa:** BCL puro (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`), sin Excel/COM, sin OpenXML SDK, sin ExcelDataReader. Rutas explícitas a la v3 + a las fuentes de los 4 períodos medidos (julio‑1, julio‑2, agosto‑2, octubre‑1) + a `Docs/Prueba Julio-2/Insumos/R10_Remuneracion_2026072.xlsx` y `Docs/Prueba Octubre-1/Resultados/Resultado administrativo/Remuneracion 202610-1 Total.xlsx`. Prohibido `Insumos.Raiz()` en lo nuevo. Cero escritura sobre plantillas, fuentes y archivos de evidencia. Cero código de producción en H0: solo fixture T0 + documento `Docs/T0-Plan59-Geometria-Derivada-V3.md`.

Cuatro frentes, por bytes y sin código de producción:

a. **Re-verificación de §0.4:** SHA256 de la v3, conteo de filas/cabeceras por hoja (§0.4-b), superconjunto de etiquetas vs v2 (§0.4-a), 0 celdas con `<f>` en las zonas nuevas (§0.4-c). Si un número difiere ⇒ STOP (la evidencia del plan no es la del disco).
b. **Vocabulario de cierres por hoja:** para R1/R2/R4, qué combinación (A,B,C,D,E) es cierre de grupo, de subsección, de sección y de bloque, y cuál es la fila de gran total. Incluye la confirmación de §0.4-d (trailer que precede) en las tres hojas y el caso R4 (sin dimensión de canal).
c. **Estabilidad de destinos CONSOLIDADO:** ¿`D9:D14`/`D104:D109` (y `DetRetriDestinoPorAse`) son resolubles por etiqueta (qué rótulo A/B los identifica)? Veredicto citado por celda: derivable-por-etiqueta o ancla-documentada-por-etiqueta. Ningún veredicto puede ser "literal porque es estable".
d. **2.º cuadro R2 con `A/EAB`:** mapa etiqueta→celda de los grupos `canal A` en los 5 ASE de la v3 + valores de la fuente octubre‑1 ASE2 (`+149.137,87` / `−18.613,87`); confirmación de que `R2!E285` ya no es fórmula en v3.

**Puerta H0→H1:** los cuatro frentes cerrados + fixture verde. Si algún frente demuestra que la única salida es un caso particular o un literal ⇒ STOP + escalado (no se abre H1).

| Candidata | Veredicto | Motivo |
|---|---|---|
| **M-A: T0 de cuatro frentes + derivación estructural + candado extendido + gate de corridas** | **Elegida** | Única que cumple §0.2 (elimina la clase, no la foto); precedente: T0-Plan50/52/54/55/56/57 |
| M-B: re-congelar las ventanas a la geometría v3 (`(4,121)` etc.) y re-congelar los 49+78 literales | Descartada | Re-fotografía la v3: la próxima plantilla rompe igual; viola §0.2.5 aunque "pase" hoy |
| M-C: derivar ventanas pero mantener los mapas de celdas con re-congelamiento v3 | Descartada | Deja viva la mitad de la clase (49+78 literales); el próximo cambio de fila la rompe en silencio o con `ERR-PLANTILLA` falso |
| M-D: rama por ASE/período "temporal" para cerrar julio‑2 mientras se diseña | Descartada | Caso particular prohibido por §0.1/§0.2.2; convierte la deuda en comportamiento |

### D-1. H1 — Ventanas y cabecera de bloque derivadas (decisión)

1. `RangoDatosAse` se **elimina** (no se edita, no se parametriza por versión). El bloque de cada ASE se delimita **leyendo la propia hoja**: los cierres `A=<sección>/B=Total` delimitan secciones, y la fila de gran total `A=Total` delimita el bloque; el inicio del bloque es la fila siguiente al gran total del bloque anterior (o el inicio de datos de la hoja, fijado por el T0 por hoja, nunca por ASE).
2. La identidad del bloque (qué ASE es) se verifica por el **rótulo de empresa** (sinonimia abierta `SinonimosEmpresaR1`, catálogo citado) — nunca por índice de fila. `VerificarCabeceraBloque` pasa a verificar identidad por etiqueta: el bloque contiene el vocabulario de la empresa esperada y su estructura de cierres está completa; si falta ⇒ `ERR-PLANTILLA` accionable (hoja+bloque+etiqueta esperada).
3. Debe soportar número **variable** de subsecciones por ASE y por período (la v3 ya lo demuestra: ASE3 +0 filas, ASE2/ASE4 +55). Ninguna cardinalidad de subsecciones se congela.
4. El delimitador es **uno solo parametrizado por hoja** (R1/R2/R4 tienen distinto vocabulario de cierres según el T0), nunca cinco copias por ASE.

### D-2. H2 — Destinos de valor y protecciones derivados (decisión)

1. Los mapas congelados de celdas editables y de fórmulas protegidas en R1/R2/R4 (`WorkbookLeafCellMapQ2.*`, `ProtectedFormulasPorEmpresa`) salen del camino productivo (mismo patrón LEGACY que `WorkbookLeafCellMapPorAse` del Plan 54 y `WorkbookLeafCellMapPorEmpresa` del Plan 55): quedan para fixtures históricos, ningún reader/writer productivo los consume.
2. Cada destino se resuelve por **(bloque/ASE, sección A, subsección B, ocurrencia de B en A, fuente C efectiva, canal D, concepto E, componente por encabezado)** con la maquinaria reusada —no reinventada—: `ContratoSeccionPeriodFreeR1` (`PlanificarJerarquiaR1`, `SeccionPorCierre`, `FirmaCuadroR1`, `CanalEfectivo`, `Alinear`), `SinonimosEmpresaR1` (catálogo abierto: un rótulo no tabulado solo se iguala a sí mismo), `ResolvedorConciliacionEmpresa` (Plan 55), `ResolvedorOperandosEmpresaR1` (Plan 53), identidad A/B/C/D/E de `EscritorValoresSlotR1` (Plan 56).
3. Cada protección se resuelve por la **etiqueta de cierre correspondiente**: `ValidarCeldaTieneFormula` (OpenXmlPlantillaWriter.cs:570) pasa a validar por identidad semántica —la celda encontrada es la que corresponde al (bloque, sección, cierre) exigido— en vez de comparar coordenada + fragmentos literales. Las llamadas literales (`:83-84, :102, :119, :544-547` y familia `:769-:969`) se sustituyen por resolución derivada; el fail-fast conserva nombre accionable (hoja+celda encontrada+identidad esperada).

   **3-bis (enmienda 2026-10-26, medida por el Ingeniero y confirmada por bytes).** La lista de **fragmentos es también geometría congelada**, no solo la celda. Evidencia: `R1Q2ProtectedPorAse[1] = ("F75", ["F47","F70","F16","L16","L47"])`; en la v3 esa celda es el **cierre de la subsección `Mes` de la sección `AFaseo`**, que vive en **`F89`** y su fórmula es **`=F55+F84+F16-L16-L55`** — es decir, cambiaron **los operandos además de la celda** (`F47→F55`, `F70→F84`, `L47→L55`). Consecuencias normativas:
   - La verificación compara **ROL + FORMA**, nunca la lista de operandos: el cierre de subsección debe ser una `<f>` que sume los totales de concepto de sus bloques `Mes` y reste los Especiales de su propia geometría resuelta. Número de términos, signos y a qué fila apunta cada uno se derivan del delimitador y del mapa de columnas, no se declaran.
   - Queda **prohibido** "actualizar la lista a la v3" (cambiar `("F75", [...])` por `("F89", [...])`): es M-B otra vez y caduca con la próxima versión.
   - El **mensaje de error** también deja de hablar en coordenadas: debe nombrar la identidad (`bloque ASE1 · sección AFaseo · subsección Mes · cierre`), y solo añadir la celda encontrada como contexto. Un error que se explica con un número de celda es la confesión de que seguimos pensando en filas.
   - En la v3, `F75` **no tiene ni debe tener fórmula** (es `<v>0</v>`): el mensaje actual ("debería seguir siendo fórmula") es correcto respecto del catálogo y **falso respecto de la plantilla**. Ese mensaje es exactamente el defecto que este plan elimina.
4. `DetRetriDestinoPorAse` también se deriva: el T0 (§2-D0-c) fija si `D9:D14`/`D104:D109` son resolubles por etiqueta; en cualquier caso el destino se obtiene por etiqueta (el rótulo existe aunque la fila no se mueva), nunca por literal. Si el T0 demuestra una celda sin etiqueta resoluble ⇒ STOP + decisión con el Ingeniero (no se deja el literal).
5. Template sin destino declarado ⇒ `ERR-SIN-SLOT` / `ERR-PLANTILLA` accionable (bloque+firma+sección); nunca reacomodo por posición ni remapeo silencioso. Guard anti-fórmula (`GuardEscrituraFormula`, `EscribirValorNumerico`) intacto y cruzado por todo camino de escritura.

### D-3. H3 — Conciliación por empresa en Q2 habilitada (decisión)

1. Se retira el recorte de `ProcesadorPeriodo` (omisión de la conciliación por empresa en Q2): el bloqueo (`R-MAPA-V2-R2-ASE4`, `R2!E285` fórmula en v2) **ya no existe** en la v3 (§0.5). El retiro es por verificación estructural (el T0 demuestra la celda-valor + el grupo declarado), no por edición del mapa.
2. Se verifica que el `EAB` de ASE2 (−18.613,87 del 2.º cuadro) cierra al centavo por la regla estructural (el grupo que la plantilla declara recibe el valor que la fuente trae), y se re-verifica `R-MAPA-V2-R2-ASE4` desbloqueado.
3. Si al retirar el recorte aparece un rojo de otra familia ⇒ STOP: el recorte se revierte por configuración, nunca con parche por empresa/ASE.

### D-4. Contratos (solo diseño, cero implementación en este documento)

- `DelimitadorBloquesHoja` (Infrastructure, BCL/OpenXML-lectura): dado (hoja, vocabulario de cierres del T0 por hoja) devuelve los rangos de bloque ordenados + la identidad por rótulo de empresa (vía sinonimia). Falla con `ERR-PLANTILLA` si la estructura de cierres está incompleta. Lo consumen `EscritorValoresSlotR1` (R1) y los caminos de escritura R2/R4.
- `ResolvedorDestinoValor` (Infrastructure, alimentado por la secuencia ya leída y tipada): dado (bloque, A, B, ocurrencia, C efectiva, D, E, encabezado) devuelve la celda destino o `ERR-SIN-SLOT` accionable. Cero literales de fila/celda en su implementación; las etiquetas son DATOS leídos, no código.
- `VerificadorDestinoProtegido`: dado (celda resuelta, identidad esperada) exige `<f>` con el rol estructural esperado o emite `ERR-PLANTILLA`; nunca compara fragmentos literales congelados.
- `CandadoAntiQuemadoExtendido` (tests): analizador estático sobre los caminos tocados + pruebas de injerto (celda, ventana, rama por ASE/empresa/período, mapa escondido, allow-list por empresa) que deben enrojecer.

### D-5. Candado anti-quemado extendido (decisión, dos capas)

Patrón Plans 54/55/56, extendido a la clase completa de este plan sobre `ConciliacionEmpresaCandadoAntiQuemadoTests`:

1. **Estático:** falla si en los caminos tocados (delimitación, resolución de destinos, verificación de protecciones, DetRetri) aparece: literal de celda (`F53`, `E285`, `"D73"`...), literal/rango de fila (`6`, `72`, `(6,72)`, offsets aritméticos sobre filas), `switch`/rama por ASE, literal de empresa/hoja/período/quincena fuera del catálogo abierto de sinónimos, o mapa/allow-list indexado por (empresa,ASE). Incluye **prueba de injerto que debe enrojecer** por cada forma (un `if (ase==2)`, un `"F53"`, un `(6,72)`, un mapa escondido por empresa injertados a propósito enrojecen).
2. **Conductual:** el mismo fixture resuelto en los 5 bloques produce las mismas reglas; reordenar filas dentro de la subsección coloca idéntico; una geometría desplazada (fixture con filas extra al estilo v2→v3) resuelve sin cambiar código.

Sin candado verde no hay H3.

### D-6. Gate de verificación (decisión)

- Corrida CLI real julio‑2 (`2026072 --cinco-ase`) contra `Docs/Prueba Julio-2/Insumos/R10_Remuneracion_2026072.xlsx`; octubre‑1 (`2026101`) contra `Docs/Prueba Octubre-1/Resultados/Resultado administrativo/Remuneracion 202610-1 Total.xlsx`; julio‑1 (`2026071`) como no‑regresión.
- Suite completa trx + **set-diff de rojos** contra la base de entrada: rojo fuera de lista = STOP.
- SHAs v1/v2/v3 verificados intactos antes/después de cada H; `G-FORMULA` (cero `<f>` tocada) en todo el diff.

### Architecture Validation Certificate

| Principio | Veredicto |
|---|---|
| **SRP** | ✅ Delimitador = geometría leída; Resolvedor = etiquetas→celda; Verificador = identidad semántica; Candado = juicio estático+conductual; cada pieza una sola razón de cambio |
| **OCP** | ✅ Geometría futura con la misma gramática (A/B/C/D/E + cierres) se absorbe sin código nuevo; celda nueva = etiqueta del template, nunca literal nuevo |
| **LSP/DIP** | ✅ Contratos `IPlantillaWriter`/`IValidador` intactos; campos aditivos o intactos; cambios detrás de las mismas abstracciones |
| **Best practices** | ✅ Contract-first (H0 congela vocabulario antes del juicio); fail-fast accionable (`ERR-SIN-SLOT`/`ERR-PLANTILLA` con bloque+firma+sección); tolerancia ±0,5 solo en comparación; values-only; §0.1/§0.2 (cero casos particulares, cero ramas escondidas); CRLF; sin commits |
| **Performance** | ✅ H0 acotado BCL ZIP+XML + fixtures existentes; delimitación O(filas de la hoja) una vez por corrida; regens a temp fresca; cero workbooks extra en runtime |

---

## 3. SPEC (requisitos numerados + gates no negociables)

### S0 — Gates transversales (toda H)

- G-BUILD: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- G-SUITE: suite completa trx + **set-diff de rojos vs la base de entrada**: altas/bajas una por una con causa citada; **rojo nuevo fuera de la lista H0 ⇒ STOP**. 0 `Skip`, 0 `Assert.True(true)`, 0 tests borrados, 0 tolerancias ampliadas; re-anclajes solo con evidencia.
- G-SHA: v3 `8199F864...2B1295` + v2 `0F192AD2...7F2DAB` + v1 `3086AC31...C5214` intactos antes/después de cada H.
- G-FORMULA: `<f>` byte-idéntica en todo el diff (cero `new CellFormula` fuera de `EscribirFormulaPreservando`, cero `InsertRow`/`DeleteRow`). **Cero `<f>` escritas por nosotros**.
- G-QUEMADO: cero literales de fila/celda/ventana y cero ramas por empresa/ASE/período/quincena en los caminos tocados —incluida la forma escondida en mapas y allow-lists— (grep de cierre + candado §2-D5).
- G-WORKBOOK: todo gate de negocio lee el workbook generado (texto-`<f>` + literales; prohibido `<v>` stale) o la fuente ya leída; **cero workbooks extra en runtime**; **sin Excel en el gate**.
- G-TRAZA: cada error nuevo nombra concepto + bloque + hoja + celda donde aplique (+ esperado vs obtenido donde aplique).
- G-PROC: cero commits; cero `git checkout` / `git reset`; CRLF sin mezclar en lo tocado; Excel cerrado durante las corridas.
- G-H0-READONLY: H0 no contiene código de producción (solo fixture `*T0Plan59*` + doc `Docs/T0-Plan59-Geometria-Derivada-V3.md`); el diff de `Remuneracion.Core/` + `Remuneracion.Infrastructure/` en H0 es vacío salvo esos tests.

### CA — Criterios de aceptación (vinculantes; con derivación al centavo donde aplique, CA0 del Plan 54)

- CA0: **todo número de criterio viaja con su derivación** (`fuente × etiqueta v3` + hoja/celda). Un número sin derivación no se acepta como criterio.
- **CA-B1**: `RangoDatosAse` eliminado del camino productivo; los 5 bloques de R1/R2/R4 se delimitan por cierres leídos (ASE1 R1 = fila 4 a 121 en v3, verificado por bytes, no por literal). Derivación: vocabulario de cierres H0 + corrida 2026072 sin `ERR-SIN-SLOT` por ventana.
- **CA-B2**: cero mapas congelados consumidos en producción (`WorkbookLeafCellMapQ2.*`, `ProtectedFormulasPorEmpresa`, llamadas literales de `ValidarCeldaTieneFormula` sustituidas por resolución derivada); `DetRetriDestinoPorAse` derivado por etiqueta. Derivación: candado estático verde + injertos en rojo + corrida sin `ERR-PLANTILLA` falso.
- **CA-B3**: conciliación por empresa en Q2 habilitada; `EAB` ASE2 (−18.613,87 del 2.º cuadro) con destino y cierre al centavo; `R-MAPA-V2-R2-ASE4` re-verificado desbloqueado (`E285` valor en v3). Derivación: fuente octubre‑1 × grupo `A/EAB` declarado en v3.
- **CA-B4**: candado extendido verde en sus dos capas (incluidas pruebas de injerto en rojo: celda, ventana, rama por ASE/empresa/período, mapa escondido, allow-list por empresa).
- **CA-B5**: corridas CLI `2026072 --cinco-ase`, `2026101`, `2026071` con exit 0; sin `ERR-SIN-SLOT`, `ERR-PLANTILLA`, `ERR-VALIDACION`, `ERR-INTEGRIDAD`; suite con set-diff exacto.
- **CA-B6**: plantillas byte-idénticas (v3 `8199F864...2B1295`, v2 `0F192AD2...7F2DAB`, v1 `3086AC31...C5214`).

### S1 — Criterios H0/T0 (read-only, cero código de producción)

- S1.1: SHA v3 + conteos §0.4-b + superconjunto §0.4-a + 0 `<f>` en zonas nuevas, re-medidos en disco (cualquier diferencia ⇒ STOP).
- S1.2: vocabulario de cierres por hoja (grupo/subsección/sección/bloque/gran total) citado celda por celda, con la semántica trailer de §0.4-d confirmada en las tres hojas.
- S1.3: veredicto por celda de `D9:D14`/`D104:D109` (+ `DetRetriDestinoPorAse`): derivable-por-etiqueta (con el rótulo citado) o STOP con escalado. Ningún veredicto "literal porque es estable".
- S1.4: mapa etiqueta→celda de los grupos `canal A` del 2.º cuadro R2 en los 5 ASE + valores octubre‑1 ASE2; `E285` verificada valor (no fórmula) en v3.
- S1.5: lista cerrada de Δ de suite (por defecto: ninguno — H0 es read-only; cualquier rojo nuevo en H0 ⇒ STOP). **Sin H0 verde no se abre H1.**

### S2 — Criterios H1 (ventanas y cabecera derivadas)

- S2.1: `RangoDatosAse` eliminado (el símbolo no existe en producción; grep de cierre lo demuestra); `VerificarCabeceraBloque` verifica identidad por rótulo de empresa vía sinonimia, nunca por fila.
- S2.2: bloques con subsecciones variables por ASE/período resueltos sin cardinalidad congelada (v3: ASE3 +0, ASE2/ASE4 +55 en la misma corrida).
- S2.3: CA-B1 + CA-B5(parcial H1: 2026072 sin `ERR-SIN-SLOT` por ventana) + CA-B6 reproducidos; 0 warnings.
- S2.4: **medido en H1:** de los 6 rojos heredados del rename `88b3acb`, **3 son por congelamiento de nombre** (2 `FileNotFoundException` + 1 `StartsWith` sobre nombre viejo) y se re-anclan a la familia vigente **por patrón/prefijo** (prohibido nombre exacto quemado) — hechos en H1.6; **3 son por CONTENIDO de la fuente** (`BalanceScPermutaTests` ×2, `DetRetriDesgloseTrazableTests` ×1: constantes de negocio documentadas contra los insumos anteriores) y **NO son re-anclables sin re-basar un oráculo de negocio** ⇒ quedan rojos con causa citada hasta decisión del Ingeniero. Solo fixtures: cero código de producción.

### S3 — Criterios H2 (destinos y protecciones derivados)

- S3.1: mapas congelados fuera del camino productivo (patrón LEGACY declarado; solo fixtures históricos los consumen); `ResolvedorDestinoValor` + `VerificadorDestinoProtegido` en producción.
- S3.2: guard anti-fórmula intacto y cruzado por todo camino de escritura; template sin destino ⇒ fail-fast accionable, nunca posición.
- S3.3: CA-B2 + CA-B5(parcial H2) + CA-B6 reproducidos; 0 warnings.

### S4 — Criterios H3 (Q2 por empresa + candado + regresión)

- S4.1: recorte Q2 retirado por verificación estructural; CA-B3 verde (EAB al centavo, `R-MAPA-V2-R2-ASE4` desbloqueado).
- S4.2: candado extendido verde en dos capas (CA-B4); code-review independiente **APROBADO** (0 bloqueantes).
- S4.3: suite trx + set-diff exacto vs base; CA-B5 completo (3 períodos) + CA-B6.

### S5 — Criterios H4/docs (cierre)

- S5.1: doctrina en `.opencode/project-context.md` (solo lo probado: geometría derivada, destinos derivados, Q2 por empresa habilitada, candado extendido, v3 como base única).
- S5.2: este plan como fuente de verdad; desviaciones con fecha/motivo/evidencia; cero código de producción en H4.

---

## 4. TASKS (fases H0..H4 con puertas)

### H0/F0 — T0 read-only (bloqueante; cero código de producción)

Fixture: `Remuneracion.IntegrationTests/GeometriaDerivadaV3T0Plan59Tests.cs` (BCL ZIP+XML, rutas explícitas; prohibido `Insumos.Raiz()` en lo nuevo). Evidencia: `Docs/T0-Plan59-Geometria-Derivada-V3.md`.

| ID | Tarea | Verifica |
|---|---|---|
| T0.1 | Re-medir §0.4 en disco: SHA v3, filas/cabeceras por hoja, superconjunto de etiquetas vs v2, 0 `<f>` en zonas nuevas | S1.1 |
| T0.2 | Fijar vocabulario de cierres por hoja (R1/R2/R4) + confirmación trailer §0.4-d en las tres hojas + caso R4 sin canal | S1.2 |
| T0.3 | Veredicto por celda de destinos CONSOLIDADO/DetRetri (derivable-por-etiqueta con rótulo citado, o STOP) | S1.3 |
| T0.4 | Mapa `canal A` del 2.º cuadro R2 en 5 ASE + valores EAB octubre‑1 ASE2 + `E285` valor en v3 | S1.4 |
| T0.5 | Lista cerrada de Δ de suite + SHAs v1/v2/v3 antes/después de H0 | S1.5 |

- **Puerta H0→H1:** fixture verde + S1.1-S1.5 cerrados + APROBADO del Ingeniero. Si algún frente solo cierra con caso particular, literal o remapeo ⇒ STOP + escalado (no se abre H1).

### H1/F1 — Ventanas y cabecera derivadas (traza a §2/§3)

| ID | Tarea | Traza |
|---|---|---|
| H1.1 | `DelimitadorBloquesHoja` por hoja (cierres + gran total leídos; identidad por rótulo vía sinonimia) | D-1, S2.1 |
| H1.2 | `EscritorValoresSlotR1`: consume el delimitador; `RangoDatosAse` eliminado; `VerificarCabeceraBloque` por etiqueta | D-1, S2.1 |
| H1.3 | Extender delimitación a caminos R2/R4 (mismo componente, vocabulario por hoja del T0) | D-1, S2.2 |
| H1.4 | e2e 2026072: CA-B1 (sin `ERR-SIN-SLOT` por ventana, 5 ASE con subsecciones variables) | S2.3 |
| H1.5 | Greps de cierre G-QUEMADO (parcial H1) + SHAs + G-FORMULA + 0 warnings | S2.3 |
| H1.6 | Re-anclar los 6 rojos heredados del rename `88b3acb` a los nombres vigentes (o resolver por patrón, prohibido nombre exacto quemado). Solo fixtures de test | S2.4 |

- **Puerta H1→H2:** S2 verde + CA-B1 + 0w + SHAs + G-FORMULA + greps.

### H2/F2 — Destinos y protecciones derivados

| ID | Tarea | Traza |
|---|---|---|
| H2.1 | `ResolvedorDestinoValor` (bloque,A,B,ocurrencia,C efectiva,D,E,encabezado) con maquinaria reusada §2-D2.2; mapas congelados a LEGACY | D-2, S3.1 |
| H2.2 | `VerificadorDestinoProtegido` por identidad semántica; sustitución de llamadas literales; guard intacto | D-2, S3.2 |
| H2.3 | `DetRetriDestinoPorAse` derivado por etiqueta según veredicto T0.3 | D-2, S3.1 |
| H2.4 | e2e 2026072 + 2026071: CA-B2 (sin `ERR-PLANTILLA` falso; sin-destino ⇒ fail-fast accionable) | S3.3 |
| H2.5 | Greps de cierre G-QUEMADO (total H2) + SHAs + G-FORMULA + 0 warnings | S3.3 |

- **Puerta H2→H3:** S3 verde + CA-B2 + 0w + SHAs + G-FORMULA + greps.

### H3/F3 — Q2 por empresa + candado + regresión

| ID | Tarea | Traza |
|---|---|---|
| H3.1 | Retiro del recorte Q2 por verificación estructural (T0.4); verificación EAB ASE2 al centavo; re-verificación `R-MAPA-V2-R2-ASE4` | D-3, S4.1 |
| H3.2 | Candado extendido (estático + injertos en rojo + conductual geometría-desplazada) | D-5, S4.2 |
| H3.3 | Gate 3 períodos: 2026072 + 2026101 (vs administrativo) + 2026071 no-regresión | D-6, S4.3 |
| H3.4 | Suite trx + set-diff vs base; fuera de lista = STOP | S4.3 |
| H3.5 | Code-review independiente (`code-reviewer`, este plan como spec): **APROBADO** requerido | S4.2 |

- **Puerta H3→H4:** S4 verde + CA-B3/CA-B4/CA-B5 + review aprobado.

### H4/F4 — Docs/cierre

| ID | Tarea | Verifica |
|---|---|---|
| H4.1 | Doctrina en `.opencode/project-context.md` (solo lo probado) | S5.1 |
| H4.2 | Cierre de este plan (§8) con desviaciones fechadas si las hubo; cero código de producción | S5.2 |

**Orden sugerido:** H0 → H1 → H2 → H3 → H4. Regresión continua (suite + SHAs en cada paso). Build 0 warnings antes de marcar cada tarea completa. Sin commits (los hace el Ingeniero). Excel cerrado durante las corridas.

---

## 5. Riesgos (consolidado + gates por fase)

Ver §1.4 (R-REFOTO, R-RAMA-OCULTA, R-CIERRE-AMBIGUO, R-DETRI-FIJO, R-ORACULO-V, R-Q2-RECORTE, R-PLANTILLA).

- **Gate H0→H1:** H0 verde + S1.5 sin STOP + APROBADO del Ingeniero.
- **Gate H1→H2:** S2 + CA-B1 + 0w + SHAs + G-FORMULA + greps.
- **Gate H2→H3:** S3 + CA-B2 + 0w + SHAs + G-FORMULA + greps.
- **Gate H3→H4:** S4 + CA-B3/CA-B4/CA-B5 + review APROBADO.
- **Gate H4:** diff de docs citado; cero producción.

---

## 6. Clarification Gate (requiere aprobación del Ingeniero; sin esto no se abre H0)

1. **v3 como base única:** ¿se ratifica que la v3 (SHA `8199F864...2B1295`) es la ÚNICA plantilla vigente desde 2026-10-26 y que v1/v2 quedan históricas intactas?
2. **H0 primero:** ¿se ratifica H0 read-only bloqueante de cuatro frentes (§2-D0), con la regla "solo cierra con caso particular o literal ⇒ STOP"?
3. **H1:** ¿el bloque se delimita por cierres leídos con identidad por rótulo (sinonimia), eliminando `RangoDatosAse` en vez de re-congelarlo a la geometría v3?
4. **H2:** ¿los 49+78 literales salen del camino productivo (patrón LEGACY) y `DetRetriDestinoPorAse` se deriva por etiqueta, con STOP si el T0 demuestra una celda sin etiqueta resoluble?
5. **H3:** ¿se retira el recorte Q2 por verificación estructural (con reversión por configuración —no parche— si aparece un rojo de otra familia)?
6. **Candado y línea base:** ¿candado extendido de dos capas (estático + injerto en rojo por cada forma: celda, ventana, rama ASE/empresa/período, mapa escondido, allow-list) + set-diff con rojo fuera de lista = STOP?
7. **Alcance:** ¿Plan 58‑P2 queda explícitamente NO tocado (requiere decisión de negocio separada)?
8. **Puerta:** sin 1-7 no se abre H0; sin H0 verde no hay H1; sin candado ni review no hay H4.

Sin estas aprobaciones el plan no avanza de PROPUESTO a APROBADO, y sin APROBADO no se abre H0.

---

## 7. Residuos declarados (fuera de alcance — NO implementar)

**R-FIXTURES-INSUMOS (se cierra en H1.6):** 6 tests verdes→rojo por el rename `88b3acb` de `Docs/Prueba Julio-2/Insumos` (nombres exactos quemados en fixtures → `FileNotFoundException`). Causa externa al plan, pero es la MISMA clase de defecto (nombre congelado); se resuelve re-anclando o por patrón, solo en fixtures.

**Plan 58‑P2 (vivo, sin tocar):** columna E de disposición final vacía en R2/R4 — requiere decisión de negocio; ni derivación ni celda inventada en este plan.

**R-RECORTE-Q2 (se retira en H3, con reversión por configuración):** `ProcesadorPeriodo` omisión Q2; gate de retiro: T0.4 + CA-B3. Si H3 enrojece por otra familia ⇒ se revierte por configuración, nunca con parche.

**R-SIGMA-Q2 (vivo, no se re-litiga):** el gate Sigma no es descomposición válida del visible en Q2 (T0-Plan50 §4.2).

**R-ANCLA-EXTEMP (cerrado por Plan 57, no reabrir):** `AnclasExtempR1.cs` ya no existe; la derivación de anclas es el precedente que este plan generaliza.

**O1/O2 del Plan 56 (vivas, con gate):** candado estático parcial (`TryPlanificar`/`MensajeSinSlotJerarquia`) y registro durable del cero documentado solo en log + memoria. Este plan extiende el candado a la clase quemada (§2-D5); O1/O2 se cierran donde vivan, no acá.

**R-ETIQUETAS-SIN-DESTINO (marco aplicable, Plan 55 §7 / Plan 56 §7):** exceso futuro ⇒ error accionable + alto sin escribir; prohibido inventar celdas destino.

---

## 8. Estado

**APROBADO 2026-10-26.** Clarification Gate §6 ratificado 8/8 por el Ingeniero ("Si procede con calidad senior"). H0 habilitado (read-only).

**H0 CERRADO 2026-10-26.** 5/5 frentes PASS, **0 discrepancias** contra §0.4. Evidencia: `Docs/T0-Plan59-Geometria-Derivada-V3.md` (453 líneas) + `Remuneracion.IntegrationTests/GeometriaDerivadaV3T0Plan59Tests.cs` (1.548 líneas, 23 tests verdes). Verificación independiente del orquestador: build **0 warnings/0 errores**, producción con `git status` **vacío** (read-only respetado), SHAs v1/v2/v3 intactos.

Resultados vinculantes de H0 que H1/H2 deben consumir:
1. **36/36 celdas de destino derivables por etiqueta** (30 CONSOLIDADO + D109 + 5 DetRetri). Cero celdas sin rótulo. Regla: título de tabla en `F` + encabezado `B='ASE'/C='Concesionarios'` + `B=id` + `C=nombre largo` → columna `D`.
2. **Vocabulario de cierres (todos TRAILER)**: R1 4 niveles (canal `D/E='Total'`; empresa `C/D='Total'`; subsección `B/C='Total'`; sección `A/B='Total'`; gran total `A='Total'`), R2 3 niveles (sin subsección), R4 2 niveles (sin canal ni subsección). Trailer verificado **42/42** al centavo en R1 octubre-1. Primera fila de datos = cabecera + 2.
3. **Trampa medida**: el cuadro-resumen bajo el gran total **reutiliza firmas de cierre** (R2 `C='OPORTUNO'/D='Total'`; R4 `B='OPORTUNO'/C='Total'`). El bloque corta en el gran total; el resolvedor no casa filas del resumen.
4. **Subsecciones variables**: plantilla 5 por bloque (2 `Aplicacion` + 3 `Mes`); fuente julio-1 trae 3 (0 `Aplicacion`); octubre-1 trae 5. Siempre 3 secciones.
5. **`R2!E285` en v3** es celda vacía sin `<f>` ni `<v>` y quedó **fuera del bloque de ASE4** (cae en ASE3 196..316); el cierre real del 2.º cuadro de ASE4 es `E362`.

Decisiones ratificadas por el Ingeniero (2026-10-26):
- **D-1 = (b)**: la línea base de la suite se fija en la corrida del árbol **vigente** (1558/1518/40/0). Los 6 rojos por rename de insumos quedan citados como causa externa (`88b3acb`), no como regresión de este plan. Los insumos de Julio-2 vigentes son los de prueba.
- **D-2**: el inicio de bloque se fija por la **cabecera de bloque** (columna B = rótulo del ASE del catálogo de dominio), nunca por "fila siguiente al gran total anterior". Coherente con la norma de §0.1: se compone por conceptos (ASE → sección → subsección → fuente → canal → concepto → componente), no por fila.

Puertas vigentes: sin H0 verde no hay H1; sin candado ni review no hay H4. Padre: Plan 58. Madre doctrina: `.opencode/project-context.md` + Plans 55/56/57 (+ 40/31 para el oráculo).

---

## 9. H1 — estado y evidencia (2026-10-26)

**H1/F1 EJECUTADO. `RangoDatosAse` ELIMINADO (0 ocurrencias en producción).** Verificación independiente del orquestador: build **0 warnings/0 errores**; fixture `DelimitadorBloquesHojaH1Plan59Tests` **8/8 verde**; `git status` de producción = solo los 3 archivos del camino tocado (`DelimitadorBloquesHoja.cs` nuevo, `LectorTextoHoja.cs` nuevo, `EscritorValoresSlotR1.cs` modificado); plantillas e insumos **intactos** (v3 `8199F864…2B1295`).

**Diseño entregado:**
- `DelimitadorBloquesHoja`: único, parametrizado por hoja (R1 4 niveles / R2 3 / R4 2), gramática de cierres **como datos**; inicio = rótulo del ASE del catálogo de dominio (D‑2); fin = primer `A='Total'` con B vacía; clasificación **por firma** (sin offsets); fila de encabezados localizada por estructura (≥2 textos en F..Z y cero numéricos); ambigüedad de firma ⇒ `ERR-PLANTILLA`.
- `LectorTextoHoja`: una sola regla de lectura de etiquetas (shared/inline/valor) compartida por delimitador y escritor.
- `EscritorValoresSlotR1`: consume el delimitador; identidad de bloque por rótulo vía `SinonimosEmpresaR1`; `ERR-SIN-SLOT` cita el rango delimitado.

**15/15 bloques resueltos, idénticos a H0.** Corte operativo del bloque = **gran total** (R1 ASE1: 4→86) y la **región** medida en H0 = cabecera→(siguiente cabecera−1) (4→121); ambas se asertan en el fixture y la región **no** se usa como geometría de producción.

**Corrida real 2026072 --cinco-ase con v3:** `ERR-SIN-SLOT` = **0 ocurrencias**. Primer error alcanzado: `ERR-PLANTILLA` en `Reporte Componentes R1!F75` ("debería seguir siendo fórmula; se detectó un valor fijo") — en la v3 esa celda es `<v>0</v>` **sin `<f>`**. Es **deuda H2 declarada**, no regresión: el catálogo congelado está equivocado respecto de la v3.

**Suite:** base 1558/1518/40/0 → H1 1589/1552/37/0 (+31 tests: 23 de H0 + 8 de H1); **0 rojos nuevos**; los 37 persistentes conservan mensaje byte-idéntico.

**Pendiente de decisión (no bloquea H2 arrancar, sí bloquea cerrar S2.4):** los 3 rojos de **contenido** del rename (constantes de negocio de los insumos anteriores) — re-basar con oráculo nuevo documentado o dejar rojos con causa citada.
