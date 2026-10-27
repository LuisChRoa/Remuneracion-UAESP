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
  plantilla_v4: Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx — VIGENTE desde 2026-10-27. Huella semantica (GATE, helper canonico `HuellaSemanticaPlantilla`) a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55; SHA256 de archivo 5FEDFD316980276945DF2AF29784EFFC5FE178C5CF1AA16F3C5D5F6F0A4E38E5 (705.805 bytes, dato INFORMATIVO). Es la v3 + exactamente 2 celdas de formula corregidas en `Reporte Componentes R1`: F328 `F269+F213+F141-L141-M213` -> `...-L213` y F333 `F254+F191+F126--L126-L191` -> `...-L126...`. Con la v4, 2026071, 2026072 y 2026101 dan RESULTADO OK
  plantilla_v3: HISTORICA desde 2026-10-27 (no se consume: ningun camino de produccion ni fixture de procesamiento la abre; solo la evidencia congelada `GeometriaDerivadaV3T0Plan59Tests`). Huella semantica 92935014f3611055c15f12db92c71020f741d009634e8ec59f1e4448235d72ff; SHA256 de archivo 8D0B3299C42BE16065CEB56E65A9EC329877109A3B6A30D35F6D83552B423538 (709.040 bytes). Motivo del retiro: typo de formula en R1!F328/F333 que impedia cerrar julio-1 y julio-2 (ajeno a la app)
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

### 0.3 Plantilla vigente: v4 desde 2026-10-27 (v3 histórica)

**Decisión del Ingeniero (textual, 2026-10-27):**

> "La v4 pasa a ser la vigente. La v3 será histórica, no se usará más para el procesamiento."

Archivo vigente: `Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx` — SHA256 de archivo `5FEDFD316980276945DF2AF29784EFFC5FE178C5CF1AA16F3C5D5F6F0A4E38E5` (705.805 bytes, **dato informativo**) y **huella semántica de contenido `a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55`** (§0.6). La huella es el **gate**; el SHA de archivo solo informa.

`v3` = `Docs/Plantilla estandarizada/v3/Formato Remuneracion 2026xx-x Total v3.xlsx` — queda en disco como **histórica** (huella `92935014…72ff`, SHA de archivo `8D0B3299…3538`, 709.040 bytes). Motivo: sus `R1!F328`/`R1!F333` traían el typo de fórmula que impedía cerrar julio-1 y julio-2. Ningún camino de producción ni fixture de procesamiento la abre; su única lectura es la evidencia congelada del T0 (`GeometriaDerivadaV3T0Plan59Tests`, motivo declarado en el propio fixture).

`v4` = `v3` + **exactamente dos celdas de fórmula** en `Reporte Componentes R1`:

| Celda | v3 | v4 |
|---|---|---|
| `F328` | `F269+F213+F141-L141-M213` | `F269+F213+F141-L141-L213` |
| `F333` | `F254+F191+F126--L126-L191` | `F254+F191+F126-L126-L191` |

Todo lo demás es idéntico: 40 hojas con el mismo nombre y orden, mismas `dimension`, mismo número de celdas (42.134), mismos valores. **Prueba medida** (no declarativa): `HuellaPlantillaVigenteV4Plan59Tests.V4_4` compara entrada por entrada las dos plantillas con el helper canónico y la diferencia es exactamente `{R1!F328, R1!F333}`.

v2 queda histórica (SHA `0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB`), v1 histórica (`3086AC31...C5214`). Las huellas —no los SHA— se re-verifican intactas antes/después de cada H.

**El mismo código procesa v3 y v4.** No hubo un solo cambio funcional para adoptar la v4: la app consume la plantilla como INPUT y deriva toda la geometría (H1) y todo destino/protección (H2/H3/H5) del contenido. Al re-basear los fixtures de la v3 a la v4, **todos** los asertos de geometría, delimitación y resolución por etiqueta pasaron sin editar un número ni una tolerancia. Esa es, precisamente, la prueba de que la geometría quedó derivada y no es una foto de la v3.

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

### 0.6 G-HUELLA — re-baseline 2026-10-27 (prevalece sobre cualquier SHA literal de este documento)

**Hecho medido:** los archivos `v2/...vf.xlsx` y `v3/...v3.xlsx` fueron **re-guardados en Excel** (y commiteados en `2002b3f`). El empaquetado cambió, el **contenido no**:

| Archivo | SHA de archivo al aprobar | SHA de archivo vigente | **Huella semántica (contenido)** |
|---|---|---|---|
| **v4 (VIGENTE)** | — | `5FEDFD31…E38E5` (705.805 B) | **`a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55`** |
| v3 (histórica) | `8199F864…2B1295` (705.934 B) | `8D0B3299…423538` (709.040 B) | **`92935014f3611055c15f12db92c71020f741d009634e8ec59f1e4448235d72ff`** |
| v2 (histórica) | `0F192AD2…7F2DAB` (693.400 B) | `86E48E47EF9E8FDEC2849448C2C7947CFE51310F2F85A9AAB7D1F3C1CC9D3D14` (693.983 B) | **`a4ea462b57be95938d0a25a9369bfd1e19d2d1229b75de7d89fbf71fb5991879`** |
| v1 (histórica) | `3086AC31…C5214` (692.598 B) | igual | **`cf9aa03604ed72b2252876df83abd573eb5609b14a97026ef7c1b23774490d17`** |

Las huellas de la tabla son las del **helper canónico** (`Remuneracion.IntegrationTests/HuellaSemanticaPlantilla.cs`: hojas + `dimension` + `f`/`v`/`si`/`ref` por celda, orden ordinal, SHA-256). Las cifras `c371a5fb…`/`dd7f5891…` que este documento citó en su primera versión provenían de una normalización no documentada y quedan **retiradas**: el gate es esta tabla. Nota: la v3 cambió de SHA de archivo otra vez al re-guardarse (§11 la midió `5ACE6F5A…FBF71F1`); su **huella no se movió** — exactamente el falso rojo que la doctrina evita.

Prueba de identidad v2/v3: comparación entrada por entrada de hojas, `dimension`, `<f>`, `<v>` y fórmulas compartidas — **0 diferencias de contenido** en cada re-guardado.

**Doctrina (aplica a este plan y a los gates):** el **SHA de archivo no es un gate válido** para una plantilla que un humano abre en Excel — cada "Guardar" reescribe el paquete y produce un falso rojo. El gate es la **huella semántica** (hash del contenido normalizado y ordenado); el SHA de archivo queda como dato informativo.
- Gate de entrada: **huella de la v4 `a7476ee2…8c55`**. Huella distinta ⇒ **STOP** (cambió el contenido: hojas, dimensiones, fórmulas o valores).
- Prueba de inyección del gate (`HuellaPlantillaVigenteV4Plan59Tests.V4_5`): revertir `R1!F328` a `…-M213` (el typo de la v3) sobre una copia **enrojece** — la huella de la copia no es la del gate. Un gate que no enrojece ante un cambio de contenido no es un gate.
- **Gate de la v4 = v3 + 2 fórmulas** (`V4_4`): el conjunto de diferencias medido por el helper es exactamente `{Reporte Componentes R1!F328, Reporte Componentes R1!F333}`; mismas hojas, mismas dimensiones, mismo número de celdas. Si alguien toca una tercera celda, enrojece.
- Los fixtures asertan **huella**, no SHA de archivo. Ese fue el arreglo de los 6 rojos de `G-SHA` (§10), no un re-baseline de números.

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

## 9. H1 — estado y evidencia (2026-10-26)**H1/F1 EJECUTADO. `RangoDatosAse` ELIMINADO (0 ocurrencias en producción).** Verificación independiente del orquestador: build **0 warnings/0 errores**; fixture `DelimitadorBloquesHojaH1Plan59Tests` **8/8 verde**; `git status` de producción = solo los 3 archivos del camino tocado (`DelimitadorBloquesHoja.cs` nuevo, `LectorTextoHoja.cs` nuevo, `EscritorValoresSlotR1.cs` modificado); plantillas e insumos **intactos** (v3 `8199F864…2B1295`).

**Diseño entregado:**
- `DelimitadorBloquesHoja`: único, parametrizado por hoja (R1 4 niveles / R2 3 / R4 2), gramática de cierres **como datos**; inicio = rótulo del ASE del catálogo de dominio (D‑2); fin = primer `A='Total'` con B vacía; clasificación **por firma** (sin offsets); fila de encabezados localizada por estructura (≥2 textos en F..Z y cero numéricos); ambigüedad de firma ⇒ `ERR-PLANTILLA`.
- `LectorTextoHoja`: una sola regla de lectura de etiquetas (shared/inline/valor) compartida por delimitador y escritor.
- `EscritorValoresSlotR1`: consume el delimitador; identidad de bloque por rótulo vía `SinonimosEmpresaR1`; `ERR-SIN-SLOT` cita el rango delimitado.

**15/15 bloques resueltos, idénticos a H0.** Corte operativo del bloque = **gran total** (R1 ASE1: 4→86) y la **región** medida en H0 = cabecera→(siguiente cabecera−1) (4→121); ambas se asertan en el fixture y la región **no** se usa como geometría de producción.

**Corrida real 2026072 --cinco-ase con v3:** `ERR-SIN-SLOT` = **0 ocurrencias**. Primer error alcanzado: `ERR-PLANTILLA` en `Reporte Componentes R1!F75` ("debería seguir siendo fórmula; se detectó un valor fijo") — en la v3 esa celda es `<v>0</v>` **sin `<f>`**. Es **deuda H2 declarada**, no regresión: el catálogo congelado está equivocado respecto de la v3.

**Suite:** base 1558/1518/40/0 → H1 1589/1552/37/0 (+31 tests: 23 de H0 + 8 de H1); **0 rojos nuevos**; los 37 persistentes conservan mensaje byte-idéntico.

**Pendiente de decisión (no bloquea H2 arrancar, sí bloquea cerrar S2.4):** los 3 rojos de **contenido** del rename (constantes de negocio de los insumos anteriores) — re-basar con oráculo nuevo documentado o dejar rojos con causa citada.

---

## 10. H2-bis — estado y evidencia (2026-10-26)

**Alcance:** cerrar el camino de escritura productivo con la v3 derivando el resto de coordenadas congeladas que lo bloqueaban, más el gate de huella (§0.6).

**Ejecutado:**

1. **`DetRetri`/`DetValiRetri` protegidos (A).** `WorkbookLeafCellMapQ2.DetRetriProtected` queda **LEGACY** (ningún camino productivo lo consume). Nuevo `DetallePeriodoDerivado` (Infrastructure) resuelve y verifica las 23 celdas por **ROL + FORMA**: tres familias de filas por ASE (principal · diferencia vs CONSOLIDADO · verificación de composición) identificadas por etiquetas, y forma derivada de las fórmulas que el propio CONSOLIDADO declara (expansión de hojas de origen, incluido `SUM(rango)`). Mensajes por identidad (hoja · familia · ASE · rótulo de columna). T0 por bytes: `Docs/T0-Plan59-H2bis-Detalle-Derivado.md`.
2. **G-HUELLA (B).** Helper compartido `HuellaSemanticaPlantilla` (contenido normalizado y ordenado: hojas, `dimension`, `<f>`, `<v>`, fórmulas compartidas). Las 8 aserciones de SHA de archivo pasan a huella; el SHA queda informativo. **Decisión requerida:** las huellas del §0.6 (`c371a5fb…`, `dd7f5891…`) provienen de una normalización no documentada; con el helper canónico son `92935014…` (v3), `a4ea462b…` (v2), `cf9aa036…` (v1). Se pide ratificar la tabla de `Docs/T0-Plan59-H2bis-Detalle-Derivado.md` §5.
3. **Barrido del camino de escritura (D).** `EscribirDetalleR2/R4` dejan de usar `WorkbookLeafCellMapDetalleR2R4.BloquesR2/R4` (ventanas por ASE de la v2) y delimitan el bloque con `DelimitadorBloquesHoja` (cabecera de bloque + gran total declarados). Fue el segundo bloqueo real de la corrida (`ERR-SIN-SLOT` con el bloque `3..33`).
4. **NO ejecutado — C (`REMUNERACION_*`).** Medido: esas hojas **sí** son derivables (misma gramática de tabla: título `R1 - RECAUDO OPORTUNO` en `F`, encabezado con `B='ASE'`, filas por `B=id`, cierre `C='TOTAL'`), pero su mapa histórico **pasa** con la v3 porque su geometría no se movió; queda como deuda de la siguiente tanda (junto con el resto de mapas congelados listados en el T0 §4).

**Evidencia:**

- Build solución: **0 warnings / 0 errores**.
- Suite: base 1598/1555/43 → **1600/1566/34**; **0 rojos nuevos**; **9 cerrados** (los 8 de SHA + `DivergenciaDeclaradaT0Tests.T0_CorridaVivaQ2…`). trx `Remuneracion.IntegrationTests/TestResults/p59-h2bis-suite.trx`.
- Corridas reales con la v3: `2026101 --cinco-ase` → **exit 0** (`Remuneración 202610-1 Total.xlsx`, 604.435 B, `VALIDACION_TOTAL O9=0 / P9=TRUE`); `2026072 --cinco-ase` → **exit 1** `ERR-VALIDACION` en `VALIDACION_TOTAL · D34/F34/O25` (Δ 1.101.678,21 entre la suma de las 4 hojas `REMUNERACION_*!D109` y `CONSOLIDADO!D109`): **fallo de negocio (familia H3/C)**, no de camino de escritura — el workbook se genera (`… (sin validar).xlsx`, 625.397 B).
- Integridad del `.xlsx` generado (`OpenXmlValidator` + caza de fórmulas compartidas huérfanas): **0 huérfanas** en ambos períodos; único hallazgo `pageSetup/@verticalDpi=0`, **presente ya en la plantilla v3** (artefacto heredado, ajeno al P0 del Plan 58).
- Plantillas intactas (SHA de archivo): v3 `5ACE6F5A…FBF71F1`, v2 `86E48E47…9D3D14`, v1 `3086AC31…A49C5214`.
- Tests nuevos: `DetalleDerivadoH2bisPlan59Tests` (verde en la base preparada del período + rojo al injertar un valor en una celda protegida, con mensaje por identidad).

**Abierto:** (i) ratificar las huellas de `G-HUELLA`; (ii) C (`REMUNERACION_*`) y el resto del barrido D; (iii) `ERR-VALIDACION` de `2026072` (composición de las `REMUNERACION_*` por empresa vs CONSOLIDADO) — requiere decisión de negocio (H3).


---

## 11. Fix del asiento DetRetri (2026-10-26) — cerrado

**Causa raíz medida:** OpenXmlPlantillaWriter.cs:346 escribía el asiento DetRetri solo bajo if (esQuincena2). **2026101 es quincena 1** ⇒ el asiento nunca se escribió y DetRetri!D9:D14 quedó en el literal `0` de la plantilla. La rama contradecía a `VerificadorDestinoProtegido.ValidarDetRetriEditable` (que exige D9:D14 editable en AMBAS quincenas).

**Falso verde:** el único chequeo leía `DetValiRetri D16:D20` (diferencias `D9 − CONSOLIDADO!U104`) desde el caché; con el asiento sin escribir y el CONSOLIDADO sin recalcular ⇒ `0 − 0 = 0` ⇒ `RESULTADO OK`.

**Corrección:** el asiento corre siempre, destino por etiqueta (ResolverTabla + EncabezadoRecaudoTotal + id leído), values-only. El cierre es **Σ de las filas por ASE** (no ROUND(Σ)): el R10 de 2026071 declaraba 58.210.094.822 = Σ filas, frente a 58.210.094.821 de ROUND(Σ) (1 peso de diferencia; en 2026101/2026072 ambas reglas coinciden).

**Gate nuevo (VerificarAsientoDetRetri):** relee el .xlsx guardado y compara lo escrito contra el dominio; **0 escrito donde el dominio declara monto ⇒ `ERR-ESCRITURA` y no se entrega**. Prueba de inyección forzando ceros: enrojece.

**Evidencia:** 2026101 D9:D14 = 14.842.425.136 / 11.302.637.605 / 8.823.787.300 / 10.483.172.762 / 7.550.668.665 / 53.002.691.468 (6/6 idénticos al oráculo administrativo). 2026071 6/6 contra el R10 oficial. 2026072 6/6 en el (sin validar). ~~su único rojo restante es el typo de plantilla `R1!F328` (-M213 en vez de -L213), que es de la propietaria~~ — **el typo quedó cerrado por la v4 el 2026-10-27** (`R1!F328` = `…-L213`; §14): 2026072 cierra con `RESULTADO OK` y sus 6 celdas son 17.450.228.673 / 20.516.143.970 / 15.221.896.467 / 7.179.595.396 / 12.073.344.662 / 72.441.209.168.

**Pendientes declarados (decisión del Ingeniero):**
1. `DetValiRetri!D9:D14` (tarifas TCS/TLU/TBL/TRT) queda en 0: el oráculo los llena, pero su fuente (tarifas del CONSOLIDADO U..AG) no está en el dominio del período.
2. El desglose J/L (DetRetri) e I/O (DetValiRetri) sigue Q2-only; medido que en Q1 el oráculo usa N(1ª) (`InterventoriaDeclarada.PrimeraQuincenaPorAse`) ⇒ hay valor para eliminar la rama sin quemar nada.
3. Modo 1-ASE nunca escribe DetRetri (no puebla DetRetriQ2): hueco preexistente.

**Nota de gate:** la plantilla se re-guardó en Excel otra vez (SHA de archivo `8D0B3299…`); contenido idéntico (0 diferencias sobre el dump normalizado). El gate válido sigue siendo la **huella semántica**, no el SHA de archivo.

---

## 12. Hueco del bloque R2/R4 en Q1 — cerrado (2026-10-26)

**Causa raíz medida:** `ExcelDataReaderWorkbookLeafInputReader` construía `DetalleR2`/`DetalleR4` solo bajo `if (esQuincena2)`. En la primera quincena el bloque `Rem. Anticipos R2` / `Reversion Pagos R4` quedaba entero en los literales `0` de la plantilla en sus tres niveles —detalle por canal/concepto, totales por canal/empresa y cierre de sección—: el cierre de cuadro (`E57 = E25+E42-K25`) daba 0, `CONSOLIDADO!D28:D32` se caía y el archivo se contradecía a sí mismo (el `DetRetri`, escrito desde el dominio, declaraba monto). Ningún gate de forma lo veía: la forma siempre fue correcta (literales, sin `<f>`).

**Corrección (producción):**
1. La rama se **elimina**: el detalle de R2/R4 es del PERÍODO y se construye en ambas quincenas. La única dependencia legítima de quincena en ese reader (`LeerRecaudosEmpresa`: par `D/E` en Q1 vs `F/G` en Q2 del `RESUMEN MES`) se conserva — es selección de columnas de la fuente, no del algoritmo.
2. **Gate nuevo** `OpenXmlPlantillaWriter.VerificarDetalleR2R4` (corre junto a `VerificarAsientoDetRetri`, sobre el `.xlsx` ya guardado): por ASE, por **nivel** del bloque y por **componente** (encabezado leído de la hoja) exige que lo escrito sea lo que la FUENTE declara para ese nivel; un nivel en cero o desalineado aborta con `ERR-ESCRITURA` y el parcial se borra. Los niveles se clasifican con la MISMA gramática de cierres de la hoja (`DelimitadorBloquesHoja.Clasificar`, nuevo accesor), aplicada a la fuente y a la hoja. Cero literales de fila/celda y cero ramas por ASE/empresa/período (el bloque se delimita por cabecera de ASE + gran total declarados; las columnas salen de la fila de encabezados).
   - **Medido y decidido:** no se exige «la cadena de agregación debe cerrar». En agosto-2 · ASE4 la propia fuente publica su cierre de sección sin que sea la suma de sus cierres por empresa (filas de detalle fuera de todo grupo en el cuadro `Subs/Cont`). Exigir esa identidad sería inventar una regla de negocio y enrojecería un entregable correcto.
3. **Prueba de inyección** (`DetalleR2R4EscritoPlan59Tests.INYECCION_...`): vaciar solo el nivel de detalle de ASE2 en R2 y, en otra pasada, el bloque entero de ASE2 en R4 ⇒ `ERR-ESCRITURA` en ambos casos.
4. **Contraste contra el oráculo administrativo** (`Q1_2026101_...`): los 5 ASE × R2/R4 coinciden componente por componente con `Docs/Prueba Octubre-1/Resultados/Resultado administrativo/Remuneracion 202610-1 Total.xlsx`. La única diferencia es la partición entre `Deb/Cred` y `Deb/Cred-otros`: el administrativo mueve a «otros» un subconjunto que la fuente trae en `Deb/Cred` (ASE2 · `Vlr Intereses` de OCCIDENTE, 6.759.294,81). Esa clasificación no está en la fuente; se verifica el TOTAL del par (coincide al centavo) y queda declarada como pendiente de negocio, misma familia que Plan 58-P2.

**Bases de fixture migradas a la v3 (la v2 no declara el grupo `canal A / empresa EAB` del SEGUNDO cuadro de R2 que la v3 agregó en los 5 ASE):**

| Fixture (archivo) | Base antes | Base después | Por qué la expectativa se conserva |
|---|---|---|---|
| `Insumos.cs` · `PlantillaAgosto2026082` (consumida por `DetalleR2R4EscrituraTests`, `ValidaControlRecaudoF10Tests`, `IntegridadCalcChainYSelloFechasTests`, `R1Q2ResolucionPorFirmaTests`, `ValidadorTotalesR1GateTests`, `BalanceScPermutaTests`, `DetRetriDesgloseTrazableTests`) | base `2026082` preparada desde la v2 | base `2026082` desde la **v3** (`BaseDelPeriodoV3`) | Fuente agosto-2 · ASE4: el 2.º cuadro trae la fila `Subsidio(-)/Contribucion(+)` del grupo `A/EAB` (−12.462,26) y sus cierres. Con la v2 esa fila no tenía slot y el detalle escrito sumaba 111.226.162,61 frente a los 111.213.700,35 de la fuente: **el bloque salía incompleto**. Con la v3 los tres niveles cuadran con la fuente (el gate lo exige) |
| `AnclasDerivadasH2Plan57Tests.cs` · `PlantillaOficial()` | v2 | **v3** | La corrida de 2026101 sobre v2 muere con `ERR-SIN-SLOT` (ASE2, canal `A/EAB`, −18.613,87 del 2.º cuadro): la v2 no declara ese grupo. Con la v3 la corrida cierra (exit 0) y las anclas siguen derivándose por etiqueta — los 23 casos del fixture quedan verdes |
| `AnclasDerivadasH2Plan57Tests.H2_1_ConductualFormulaInjertada...` | corrida sobre `2026072` | corrida sobre `2026101` | El conductor (injerto de un término falso ⇒ el núcleo dispara con el Δ exacto) necesita una corrida cuyo resultado exista. `2026072` no cierra con la v3 por el defecto de la plantilla `R1!F328` (`M213` donde debe ser `L213`), ajeno a la app y ya declarado; `2026101` cierra con exit 0 y el aserto no depende del período |
| `DetalleR2R4EscrituraTests.cs` · localización del slot | `WorkbookLeafCellMapDetalleR2R4.BloquesR2/R4` (ventana v2) | `DelimitadorBloquesHoja.DelimitarBloque` (cabecera + gran total) | Con la v3 la ventana congelada de la v2 ya no es el bloque; el slot se busca dentro del bloque REAL. El aserto es el mismo (primer `Vlr Servicio` == fuente) |
| `DetalleR2R4Tests.LeerLeafInputs_Q1_NoPueblaDetalle` | aserto de la rama (Q1 sin detalle) | `LeerLeafInputs_Q1_PueblaDetalleIgualQueQ2_SinAlterarAgregados` | El aserto anterior **codificaba el defecto**. Lo respalda el oráculo: 2026101 y el R10 de 2026071 (ambas primera quincena) traen el detalle lleno |
| `LecturaR1SinRamaQuincenaTests` (allow-list + capa conductual) | cupo 2 de `var esQuincena2 = ...`; `Assert.Null(q1.DetalleR2)` | cupo 1 (solo `LeerRecaudosEmpresa`); Q1 == Q2 fila por fila | El candado se **aprieta** (una ocurrencia menos que declarar) y la identidad Q1≡Q2 se vuelve más fuerte: antes solo comparaba R1, ahora también la secuencia completa del detalle |

**Evidencia:**
- Build solución: **0 warnings / 0 errores**.
- Suite: base 1605/1571/34 → **1608/1574/34** (+3 tests nuevos). **0 rojos nuevos, 0 cerrados, 0 rojos desaparecidos**; el único nombre que sale del listado es el rename citado en la tabla (aserto del defecto → aserto del comportamiento correcto). trx `TestResults/p59-frente3-final.trx`.
- Corridas reales con la v3: `2026101` → **exit 0** (`Remuneración 202610-1 Total.xlsx`, 618.867 B); `2026071` → exit 1 `ERR-VALIDACION` (Δ 1.361.600,46 = el defecto `R1!F328` de la plantilla); `2026072` → exit 1 `ERR-VALIDACION` (Δ 1.101.678,21, **idéntico** al medido en §10). Los tres veredictos se compararon **con la rama restaurada** (estado pre-fix) sobre la misma base v3: `58208733220.04 != 58210094820.50` y `72440107489.50 != 72441209167.71` en ambos estados ⇒ la corrección cambia el CONTENIDO del entregable, no el veredicto de validación (que arrastra el typo de la plantilla). Cero `ERR-SIN-SLOT`, `ERR-PLANTILLA`, `ERR-ESCRITURA`, `ERR-INTEGRIDAD` en las tres.
- Bloque R2 · ASE1 · 2026101, oráculo vs generado (misma semántica, distinta geometría de filas): E(canal E) 28.281.218,77 = 28.281.218,77; E(H) 604.175,43 = 604.175,43; E(O) 448.577,54 = 448.577,54; E(T) 1.942.770,68 = 1.942.770,68; cierre `Componente` 31.276.742,42 = 31.276.742,42; cierre `Subs/Cont` −688.671,40 = −688.671,40; gran total 30.588.071,02 = 30.588.071,02. Filas de detalle con valor: 12 = 12 (ASE1); 77 = 77 en los 5 ASE.
- Plantillas intactas: `git status` sin cambios en `Docs/**/*.xlsx`; SHA de archivo v3 `8D0B3299…`, v2 `86E48E47…` (los mismos de §11/§0.6).

**Residuos declarados (no tocados):**
1. El cierre de cuadro `R2!E57` (=`E25+E42-K25`) queda con caché 0 en el archivo guardado (la app no sella esa cadena): Excel lo recalcula al abrir porque el writer deja `fullCalcOnLoad`. Los operandos ya están escritos — es cosmético y previo.
2. La partición `Deb/Cred` vs `Deb/Cred-otros` del administrativo (ver arriba) sigue sin regla derivable: pendiente de negocio.
3. ~~`Insumos.Plantilla` (2026071) y `Insumos.PlantillaQ2` (2026072) siguen apoyadas en la v2~~ — **CERRADO**: §13.4 (H5) las migró a la plantilla vigente y, desde el rebase del 2026-10-27 (§14), se preparan desde la **v4** (`Insumos.BaseDelPeriodoV4`). Hoy no queda ninguna base de fixture de procesamiento sobre v3.

---

## 13. H5 — erradicación de literales de período (2026-10-26) — cerrado

**Mandato del Ingeniero (textual):**

> "¿La app tiene reglas o cosas quemadas por período? Si es período tal haga algo… Eso es inadmisible, eso no lo permitimos. Sería pésimo si lo llegáramos a tener y no puede quedar en función de un período."

**Doctrina aplicada:** el período es un DATO del dominio (`Periodo.CodigoCompleto`). El nombre de las hojas que lo llevan se LEE del workbook que se está procesando (`HojasPeriodo`). Un literal de período en código —aunque hoy "coincida"— es una regla por período esperando a romper: es la misma clase de defecto que `RangoDatosAse`, no un caso menor.

### 13.1 Auditoría medida (inventario ANTES, código de producción)

Se escanearon `Remuneracion.Core`, `Remuneracion.Infrastructure`, `Remuneracion.WinForms` y `Remuneracion.Cli`, juzgando **código** (los comentarios y la documentación XML se descartan: citar un período como evidencia no es quemarlo).

| Archivo:línea | Contenido | Veredicto |
|---|---|---|
| `Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs:47` | `private const string AnclaSufijoHojasQ1 = "2026071";` | **Quemado.** Ancla de naming escrita a mano; `ReemplazarSufijo` la sustituía por `periodo.CodigoCompleto`. Con la v3 el sufijo real de la plantilla es otro: el ancla era una foto. |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCellMapBalanceSc.cs:80-84, 96, 146-173` | `("DetRetri2026071", "J32", […])`, `(HojaBce, "H3", ["DetRetri2026071", "J9"])`, `("DetValiRetri2026071", …)` | **Quemado.** La foto de protecciones de la v1 codificaba el nombre de hoja de UN período concreto en sus fragmentos. |
| `Remuneracion.Core/Models/DivergenciasDeclaradasCatalogo.cs:111, 141` | `"… oficial 2026101…"`, `"… (2026072 desde el oficial) …"` | **Prosa de evidencia con literal.** Se sustituyó por el nombre del período (`octubre-1`, `julio-2`); ninguna aserción dependía del literal (los fixtures del catálogo solo exigen no-vacío, `"2 filas"`, `"Plan 46"` y rutas de documentos). |
| `Remuneracion.Cli/OpcionesCli.cs:219` | `--periodo AAAAMM[Q]  Período (p. ej. 2026071)…` | **Legítimo, declarado exento.** Texto de ayuda del CLI: documenta el FORMATO del argumento; no decide hoja, celda, rama ni valor. |

Resto de coincidencias (`20\d{4,6}` en 30+ archivos): **todas en comentarios/documentación XML** (`ExcelDataReaderWorkbookLeafInputReader.cs:58`, `WorkbookLeafCellMapQ2.cs:8-24`, `OpenXmlPlantillaWriter.cs:24-1622`, `WorkbookLeafModel*`, `Periodo.cs`, `NombresHojaPeriodo.cs`, `Form1.cs:797`, …). No se tocan: son evidencia fechada, no comportamiento.

### 13.2 Punto 1 — el sufijo de hoja se DERIVA de la plantilla

Nuevo `Remuneracion.Infrastructure/Excel/HojasPeriodo.cs`:

- `DetRetri(parte, periodo)` / `DetValiRetri(parte, periodo)` enumeran las hojas del workbook abierto, identifican la raíz por **estructura** (`DetRetri`/`DetValiRetri` + 7 dígitos) y **extraen el sufijo que la plantilla declara**.
- El nombre resuelto se compone **sustituyendo ese sufijo** por `periodo.CodigoCompleto`. En una base preparada (Phase 0 renombra las tres hojas del período) la sustitución es identidad; en la plantilla cruda (v3 declara su propio sufijo) resuelve el nombre real.
- Si el workbook no declara exactamente una hoja de esa raíz, o el período pedido no está presente, **fail-fast `ERR-PLANTILLA`** nombrando sufijo declarado, período en proceso y hoja esperada. La app no escribe en la hoja de otro período.
- `AnclaSufijoHojasQ1` + `ReemplazarSufijo` + `ProtegidasBceParaPeriodo` **eliminados**. `VerificadorDestinoProtegido.ValidarDetRetriEditable` y `DetallePeriodoDerivado.Validar` consumen la hoja REAL del período (antes componían el nombre desde el dominio: mismo resultado, pero sin verificar que la base sea la del período).

### 13.3 Punto 2 — la foto de protecciones BCE sale del camino productivo

- `WorkbookLeafCellMapBalanceSc.Protegidas` (78 celdas + fragmentos con nombre de hoja de período) **se retira de producción**. Se conserva como evidencia histórica del T0 en `Remuneracion.IntegrationTests/BalanceScGoldenV1.cs` (patrón LEGACY de los demás mapas: solo lo consumen fixtures de catalogación).
- Nuevo `VerificadorBalanceScDerivado` (producción), que corre en el tronco común de ambos caminos de escritura (multi-ASE y single-ASE, Q1 y Q2) y deriva TODO de lo que la hoja declara:
  1. **Bandas por rótulo**: `ASE · SUBSIDIO · CONTRIBUCION · TOTAL BSC · SISTEMA` y `ASE · REMUN TOTAL · VALIDA 1 · VALIDA 2 · DIFERENCIA VALIDA 2`; las filas, por el **id de ASE** que la propia hoja numera (5 del catálogo, sin duplicados).
  2. **Vínculo con el detalle del período**: `SISTEMA` de cada ASE = balance del ASE en la hoja `DetRetri` **real** (columna y celda resueltas por el rótulo `BALANCE DE SUBSIDIO O CONTRIBUCIONES` + id de ASE), y el agregado contrastado contra la **fila de cierre** de esa tabla.
  3. **Vínculo con el CONSOLIDADO**: `REMUN TOTAL` = celda que la tabla `REPORTE CONSOLIDADO TOTAL DE RECAUDO` declara para ese ASE (por encabezado).
  4. **Sumas y cuadres**: Σ de cada columna de valor sobre su tramo de ASE, Σ cruzada `SUBSIDIO+CONTRIBUCION`, diferencia por ASE y por agregado, cuadre `ROUND` entre ambas sumas, y las comparaciones `VALIDA 1/2` de ASE y de cierre.
  5. El rol **diferencia** no tiene rótulo en la plantilla: se deriva por **forma** (`TOTAL BSC − SISTEMA`) y debe ser una sola columna en toda la banda.
- Cero literales de fila/celda/período. Mensajes `ERR-PLANTILLA` por identidad (banda · rol · ASE · rótulo).
- **El camino single-ASE deja de usar su foto v1** (`WorkbookLeafCellMap.ProtectedFormulas`: F75/F77/E47/D73, que con la v3 ya no son fórmulas) y pasa por el mismo verificador derivado: una sola definición de "protegido" para los dos caminos.
- **El gate sigue vivo (inyección, `BalanceScDerivadoH5Plan59Tests`, 4/4 verde):** base intacta ⇒ pasa; `SISTEMA` convertido en valor ⇒ `ERR-PLANTILLA` nombrando `SISTEMA`/`DetRetri`/celda; Σ de tramo convertida en valor ⇒ `ERR-PLANTILLA` nombrando la suma de `TOTAL BSC`; `TOTAL BSC` con un término de más ⇒ `ERR-PLANTILLA` nombrando el rol.

### 13.4 Punto 3 — fixtures y bases

`Insumos.Plantilla` (`2026071`) y `Insumos.PlantillaQ2` (`2026072`) pasan a la **v3 vigente** (`BaseDelPeriodoV3`). Por qué la expectativa se conserva: la v3 es **superconjunto estricto** de la v2 en etiquetas (§0.4-a: 0 faltantes en las tres hojas), conserva las 40 hojas y las columnas, y **no mueve la geometría de escritura del CONSOLIDADO** (§0.4-c: D9:D13, D28:D32, D47:D51, D66:D70, D85:D89, D104:D109, 1.993 fórmulas re-apuntadas con la misma forma). Lo que cambia es la FILA de los cierres de bloque — exactamente la clase de defecto que H1/H2 eliminaron.

Fixtures que quedan **explícitamente** sobre la v2 (`Insumos.BaseV2DelPeriodo`, motor declarado en el propio helper). No se relaja ni se borra una sola aserción; lo que cambia es la base contra la que se ejerce, y el motivo es siempre uno de estos dos:

**(a) su OBJETO es la geometría v2** (caracterización de una foto LEGACY):

| Fixture | Base antes | Base después | Motivo |
|---|---|---|---|
| `MapaQ2T0Tests` (7 casos rojos evitados) | `Insumos.PlantillaQ2` (v2) | `Insumos.BaseV2DelPeriodo("2026072")` | Su sujeto ES el mapa congelado `WorkbookLeafCellMapQ2` (F75/E47/D32/F330…), LEGACY. Ese mapa describe la v2; con la v3 sus celdas se movieron (§0.4-b) y "actualizar la lista" es la re-foto que §0.2.5 prohíbe. La cobertura vigente la dan `DetalleR2R4EscritoPlan59Tests` / `DestinoValorDerivadoH2Plan59Tests` (v3, por identidad). |
| `DetalleR2R4EscrituraTests` (casos de julio) | `Insumos.PlantillaQ2` (v2) | `Insumos.BaseV2DelPeriodo("2026072")` | Los valores citados son los de las FILAS de la evidencia T0b del Plan 29 (R2 `E3/E4/E7/E9`, `E78..E80`, `E408..E425`, bloque `67..124`). Con la v3 los bloques de R2 son `1/75/196/317/441` y esas cifras viven en otras filas. La cobertura v3 equivalente (localización por firma+canal, 5 ASE, ambas quincenas) está en `DetalleR2R4EscritoPlan59Tests` y `DetalleR2R4EscritoPlan59Tests.INYECCION_*`. |
| `GoldenDetRetriQ2Tests` | `Insumos.PlantillaQ2` (v2) | `Insumos.BaseV2DelPeriodo("2026072")` | Es la caracterización del **caché golden** de julio-2 (manual del administrativo, geometría v2) y usa `WorkbookLeafCellMapQ2.ObtenerDetRetriDestino` (LEGACY). Sobre la v3 el golden celda-a-celda deja de ser oráculo: su geometría se movió. |

**(b) motivos v3→v2 que la v4 dejó sin efecto (histórico del H5).** En H5 estas corridas se anclaron a la v2 porque sobre la **v3** no podían cerrar: `R1!F328` escribía `-M213` donde debe ser `-L213` (y `R1!F333` traía `--L126`). **La v4 corrigió ambas celdas el 2026-10-27** (§14) y el motivo desapareció; los fixtures se dejan sobre la v2 tal como los fijó H5 (moverlos otra vez solo cambiaría la base sin cambiar su objeto) y su migración a la v4 es decisión del Ingeniero — ver §14.5.

| Fixture | Base antes | Base después | Motivo |
|---|---|---|---|
| `ProcesadorPeriodoTests.Ejecutar_Periodo2026072_ConLectorOracle…` | `Insumos.PlantillaQ2` (v2) | `Insumos.BaseV2DelPeriodo("2026072")` | Su objeto es el lector-oráculo sobre una plantilla con caché 0; julio-2 sobre la v3 no cierra (Δ 1.101.678,21 = `R1!F328`) y no se puede exigir `exit 0` donde el plan exige `ERR-VALIDACION`. |
| `SelladoCacheFormulaT2Tests` (flujo Q2) | `Insumos.PlantillaQ2` (v2) | `Insumos.BaseV2DelPeriodo("2026072")` | El sellado por read-back exige el valor EVALUADO (`1`); con el typo de `R1!F328` los sub-bloques de julio-2 evalúan `FALSE` y el sellado devuelve `0`. Un `0` sellado en una plantilla defectuosa sería mentir en el archivo. |
| `ParidadCliTests.Paridad_5Ase_Q1`, `…_Q2`, `CodigosSalidaCliTests.Ejecucion_5Ase_Q1` | `Insumos.Plantilla` / `PlantillaQ2` | `Insumos.BaseV2DelPeriodo(…)` | **Ya estaban rojos en la base** (`p59-frente3-final.trx`): julio-1 arrastra un Δ de 22.921,61 en `VALIDACION_TOTAL` y la paridad Q2 un desvío de 13.084.233.937,44 en una celda del mapa v2. Re-anclarlos a la v2 restaura **el mismo fallo, con los mismos nombres** (verificado: misma aserción, mismo Δ). No se relajan: se les devuelve su base histórica para que el rojo no cambie de causa. |

Resultado: **set-diff exacto** contra la base — ver §13.6.

### 13.5 Punto 4 — candado nuevo (`CandadoPeriodoQuemadoPlan59Tests`, 6/6 verde)

Extiende el patrón de `ConciliacionEmpresaCandadoAntiQuemadoTests` a la clase de defecto "período quemado":

- **Estático**: falla si en los cuatro proyectos de producción aparece (1) un **código de período** (`20` + 4..6 dígitos, delimitado) o (2) un **nombre de hoja con sufijo de período** (`DetRetri…`, `DetValiRetri…`, `Informe AFaseo Recaudo 20…`). Juzga el texto sin comentarios **línea a línea**, no solo lo entrecomillado: así ve también los literales crudos multilínea (donde vive el texto de ayuda del CLI).
- **Lista exenta CERRADA con motivo**: única entrada = `Remuneracion.Cli/OpcionesCli.cs` · `p. ej. 2026071` · "texto de AYUDA que documenta el FORMATO del argumento; no decide comportamiento ni selecciona hoja, celda ni rama". La prueba `H5_LaExcepcionDeclarada_CubreSoloElTextoDeAyuda` verifica que la misma frase en otro archivo **sí** enrojece.
- **Injertos que deben enrojecer** (verde en el candado = el candado no es decorativo): `"DetRetri" + "2026071"`, `if (periodo.CodigoCompleto == "2026101")`, `var ancla = "2026072"`, `"DetRetri2026071"`, `"DetValiRetri2026101"`, `"Informe AFaseo Recaudo 202607-2"`.
- **Anti-regresión del LEGACY**: falla si el nombre `BalanceScGoldenV1` reaparece en producción (que la foto no vuelva a ser fuente de coordenadas).
- **No enrojece con lo legítimo**: comentarios/documentación XML que citan períodos reales, `periodo.CodigoCompleto`, `HojasPeriodo.DetRetri(...)`, `<AAAAMM|AAAAMMQ>`.

### 13.6 Evidencia

- **Build**: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- **Suite completa**: `1618/1584/34` (10 tests nuevos: 6 candado + 4 BCE) contra la base `1608/1574/34`. **Set-diff exacto: 34 = 34, cero rojos nuevos, cero rojos cerrados, cero tests borrados, cero tolerancias ampliadas.** trx: `Remuneracion.IntegrationTests/TestResults/p59-h5-cierre.trx`.
- **Corridas reales (v3, `--cinco-ase`)**:
  - `2026101` → **exit 0** `RESULTADO OK` (`Remuneración 202610-1 Total.xlsx`, **618.867 B**, idéntico al de §12 ⇒ el entregable no cambió).
  - `2026071` → exit 1 `ERR-VALIDACION`, Δ **1.361.600,46** (`R1!F328`, propietaria) — el mismo de §12.
  - `2026072` → exit 1 `ERR-VALIDACION`, Δ **1.101.678,21** (`R1!F328`) — el mismo de §12.
  - Las tres: **0 `ERR-SIN-SLOT`, 0 `ERR-PLANTILLA`, 0 `ERR-ESCRITURA`, 0 `ERR-INTEGRIDAD`**.
- **Asiento y desglose intactos**: `2026101` sigue con `DetRetri!D9:D14` 6/6 contra el oráculo administrativo (14.842.425.136 / 11.302.637.605 / 8.823.787.300 / 10.483.172.762 / 7.550.668.665 / 53.002.691.468) y `VALIDACION_TOTAL` de octubre `O9=0 / P9=TRUE`; el desglose J/L/I/O y el detalle R2/R4 de ambas quincenas no se tocaron en H5 (los gates de §11/§12 son los mismos y siguen verdes en la suite).
- **Plantillas intactas**: SHA de archivo v3 `8D0B3299C42BE16065CEB56E65A9EC329877109A3B6A30D35F6D83552B423538` (el de §12), v2 `86E48E47…9D3D14`, v1 `3086AC31…A49C5214`; `git status` **sin cambios en `Docs/**/*.xlsx`**.
- **Greps de cierre**: inventario DESPUÉS en producción = **0 códigos de período** y **0 nombres de hoja con sufijo**, fuera de la exención declarada. `HojasPeriodo` es la única fuente del sufijo; `WorkbookLeafCellMapBalanceSc` ya no contiene celdas ni nombres de período.
- **No tocado en H5 (mandato)**: el typo `R1!F328` —**cerrado por la v4 el 2026-10-27**, ver §14— y la clasificación `Deb/Cred` vs `Deb/Cred-otros` (regla pendiente de la administradora, §12).

### 13.7 Decisiones que quedan al Ingeniero

1. ~~**`R1!F328` de la v3**: julio-1 y julio-2 no cierran por ese typo (Δ 1.361.600,46 / 1.101.678,21). Mientras no se corrija en la plantilla, todo fixture que exija "cierre de julio sobre la v3" es irrealizable por diseño.~~ **CERRADO 2026-10-27**: la v4 corrige `R1!F328` (`…-L213`) y `R1!F333` (`…-L126…`); julio-1 y julio-2 cierran con `RESULTADO OK` (§14.3). La afirmación histórica sobre la v3 sigue siendo cierta —por eso la v3 pasa a histórica— y ningún fixture vigente depende ya de ese defecto.
2. **Reducción declarada de cobertura en las columnas derivadas del detalle** (`DetRetri!J/K/M`, `DetValiRetri!M/O`): la foto vieja exigía la PRESENCIA del fragmento `BCESCPORFACT.!F3` en esas fórmulas; el verificador derivado hoy cubre el detalle del período por su primera columna de valor (`DetallePeriodoDerivado`, §10) y el vínculo inverso BCE→detalle por identidad. Extender la derivación a las columnas de composición heterogénea (BCE/INTERVENTORIA/sumas anidadas) es trabajo del siguiente frente, no una re-foto.
3. **`WorkbookLeafCellMapBalanceSc.EditablesPorAse`** (D/E por ASE: `E3/D3`…), último mapa de coordenadas vivo en esa clase y objetivo de escritura del balance. Es el resto del "barrido D" de §10: derivable por (id de ASE + encabezado) con la misma maquinaria que ya usa `VerificadorBalanceScDerivado`; queda declarado, no silenciado.
4. **`Docs/T0` de H5**: no se generó documento T0 porque la medición de este frente se hizo sobre código y plantilla ya medidas en §0.4/§0.6; toda la evidencia vive en este §13 y en los fixtures citados.
---

## 14. Rebase de baseline a la plantilla v4 (2026-10-27) — cerrado

### 14.1 Decisión y alcance

**Decisión del Ingeniero (textual):**

> "La v4 pasa a ser la vigente. La v3 será histórica, no se usará más para el procesamiento."

Alcance ejecutado: **re-apuntar el baseline**, no rediseñar nada. Cambió la plantilla de origen (v3 → v4) y la huella-gate; **no** cambió una sola regla de negocio, gate, criterio ni tolerancia. La v3 permanece en disco como histórica y ningún camino de producción ni fixture de procesamiento la abre.

### 14.2 Qué se re-apuntó

| Artefacto | Antes | Después |
|---|---|---|
| `Insumos.PlantillaOficialV3()` | ruta `v3/…v3.xlsx` | **`Insumos.PlantillaOficialV4()`** → `v4/…v4.xlsx` (sin alias: el nombre viejo no existe) |
| `Insumos.BaseDelPeriodoV3(periodo)` | Phase 0 desde la v3 | **`Insumos.BaseDelPeriodoV4(periodo)`** (mismo flujo, plantilla de origen v4) |
| `Insumos.Plantilla`, `Insumos.PlantillaQ2`, `Insumos.PlantillaAgosto2026082` | `BaseDelPeriodoV3` | `BaseDelPeriodoV4` |
| `AnclasDerivadasH2Plan57Tests` | `Insumos.PlantillaOficialV3()` | `Insumos.PlantillaOficialV4()` |
| `BalanceScDerivadoH5Plan59Tests` | `BaseDelPeriodoV3("2026101")` | `BaseDelPeriodoV4("2026101")` |
| `DetalleR2R4EscritoPlan59Tests` | ruta literal `v3/…` | `Insumos.PlantillaOficialV4()` (propiedad `PlantillaVigente`) |
| `DetalleDerivadoH2bisPlan59Tests` | ruta literal `v3/…` | `Insumos.PlantillaOficialV4()` |
| `DelimitadorBloquesHojaH1Plan59Tests` | ruta literal `v3/…` | `Insumos.PlantillaOficialV4()` (propiedad `PlantillaVigente`) |
| `DestinoValorDerivadoH2Plan59Tests` | ruta literal `v3/…` | `Insumos.PlantillaOficialV4()` (clase `PlantillaVigente`) |
| `HuellaSemanticaPlantilla` | solo `V3` | **`V4` (gate)** + `V3` conservada como histórica |

Renombres de test con la versión en el nombre (los demás conservan su nombre; su cuerpo ya no depende de la versión, que es el punto):
`Bases2026072_ConLaV3_…` → `Bases2026072_ConLaPlantillaVigente_…`,
`H1_2_R1_EscribeLosCincoBloquesSobreLaV3_…` → `…SobreLaPlantillaVigente_…`,
`H2_2_…NoTieneFormulaEnLaV3` → `H2_2_…NoTieneFormula`.

**Única lectura de la v3 que queda en la suite:** `GeometriaDerivadaV3T0Plan59Tests` (evidencia congelada del H0, motivo declarado en la cabecera del fixture) y `HuellaPlantillaVigenteV4Plan59Tests` (necesita el archivo histórico para probar que `v4 = v3 + 2 fórmulas`). Cero código de producción menciona la v3: la app recibe la plantilla por `--plantilla` y deriva todo de su contenido.

### 14.3 Corridas reales con la v4 (`--cinco-ase`, plantilla = ruta de la v4)

| Período | Exit | Veredicto | `DetRetri!D9:D14` (ASE1..ASE5 / total) |
|---|---|---|---|
| `2026071` | **0** | `RESULTADO OK` | 16.758.167.585 / 12.227.292.053 / 10.127.195.599 / 10.553.776.376 / 8.543.663.209 / **58.210.094.822** |
| `2026072` | **0** | `RESULTADO OK` | 17.450.228.673 / 20.516.143.970 / 15.221.896.467 / 7.179.595.396 / 12.073.344.662 / **72.441.209.168** |
| `2026101` | **0** | `RESULTADO OK` | 14.842.425.136 / 11.302.637.605 / 8.823.787.300 / 10.483.172.762 / 7.550.668.665 / **53.002.691.468** |

- `2026101` reproduce **al peso** el asiento de §11 y el total de su oráculo administrativo: la v4 no cambió nada de lo que ya cerraba.
- `2026071` cierra contra el R10 oficial (58.210.094.822) — el Δ 1.361.600,46 de §12/§13 **desapareció**: era `R1!F328`.
- `2026072` cierra con 72.441.209.168 — el Δ 1.101.678,21 de §10/§12/§13 **desapareció**: era el mismo `R1!F328`.
- Las tres: **0 `ERR-SIN-SLOT`, 0 `ERR-PLANTILLA`, 0 `ERR-ESCRITURA`, 0 `ERR-INTEGRIDAD`**; `VALIDACION_TOTAL O9 = 0; P9 = TRUE` en los 5 ASE de los 3 períodos.
- Salidas escritas en ruta temp FRESCA (`%TEMP%\p59-v4-run\<período>`), nunca sobre una salida existente y nunca sobre la plantilla.

### 14.4 Huella semántica de la v4 y prueba de inyección

| Dato | Valor |
|---|---|
| **Huella semántica v4 (GATE de entrada)** | **`a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55`** |
| SHA-256 de archivo v4 (informativo) | `5FEDFD316980276945DF2AF29784EFFC5FE178C5CF1AA16F3C5D5F6F0A4E38E5` (705.805 bytes) |
| Huella semántica v3 (histórica) | `92935014f3611055c15f12db92c71020f741d009634e8ec59f1e4448235d72ff` |
| SHA-256 de archivo v3 (informativo) | `8D0B3299C42BE16065CEB56E65A9EC329877109A3B6A30D35F6D83552B423538` (709.040 bytes) |

Calculada con el helper canónico (`HuellaSemanticaPlantilla.Calcular`, §0.6), **no a mano**: se contrastó con una reimplementación independiente (BCL ZIP+XML en PowerShell) que reproduce exactamente la huella conocida de la v3 `92935014…72ff` y da `a7476ee2…8c55` para la v4.

Fixture de identidad y gate: `Remuneracion.IntegrationTests/HuellaPlantillaVigenteV4Plan59Tests.cs` (5/5 verde).

| Test | Qué prueba |
|---|---|
| `V4_1` | El archivo vigente tiene el SHA informativo y el tamaño declarados (coincide con la medición del orquestador: 705.805 B). |
| `V4_2` | `Calcular(v4) == V4` — la huella-gate. |
| `V4_3` | La v3 sigue en disco y su huella es `92935014…72ff`; las dos huellas son distintas. |
| `V4_4` | **`v4 = v3 + exactamente 2 celdas`**: mismo conjunto/orden de hojas, mismas `dimension`, mismo número de celdas, y el conjunto de diferencias medido es `{Reporte Componentes R1!F328, Reporte Componentes R1!F333}` con las fórmulas `…-L213` y `…+F126-L126-L191`. |
| `V4_5` | **Inyección**: sobre una COPIA temporal se revierte `F328` a `…-M213` y la huella deja de ser `V4` ⇒ **el gate enrojece**. |

**Lectura del resultado:** el gate no es un número decorativo. Si mañana alguien "arregla" una tercera celda, si revierte el typo, o si el archivo pierde o gana una hoja, `V4_2`/`V4_4` fallan y la suite no deja pasar el baseline. Y como el gate es de **contenido** (§0.6), un simple "Guardar" en Excel sobre la v4 no lo rompe.

**Nota sobre `<f>`:** la única escritura de una fórmula en todo este frente ocurre dentro de `V4_5`, sobre una **copia temporal**, para probar el gate. La producción no escribe `<f>` (guard anti-fórmula intacto), no inserta ni elimina filas y ninguna plantilla fue modificada: v3 y v4 conservan byte a byte sus SHA (`8D0B3299…` y `5FEDFD31…`).


---

### 14.5 Suite completa y set-diff

Corrida: `dotnet test Remuneracion.IntegrationTests` (trx fuera del repo, `%TEMP%\p59-v4-suite3\p59-v4-suite.trx`).

| | Total | Verdes | Rojos | Skip |
|---|---|---|---|---|
| Base `TestResults/p59-h5-suite.trx` | 1617 | 1583 | 34 | 0 |
| **Con la v4** | **1622** | **1595** | **27** | 0 |

> Nota de conteo: §13.6 transcribió `1618/1584/34`; el trx en disco mide **1617/1583/34**. El set-diff se hace contra el trx (la evidencia), no contra el texto.

**Set-diff: 0 rojos nuevos · 7 rojos cerrados · 5 tests nuevos · 4 renombres.** Ninguna tolerancia ampliada, ningún gate retirado, ningún test borrado.

Rojos cerrados (los 7 comparten UNA sola causa, medida):

| Test | Causa del rojo en la base | Por qué cierra |
|---|---|---|
| `ProcesadorPeriodoTests.Ejecutar_Periodo2026071_ConLectorOracle_VeredictosValidacionesNoVacios` | la corrida Q1 sobre la v3 abortaba con `ERR-VALIDACION` (`ASE n · VALIDACION_TOTAL · D34/F34/O25`: `58208733220.04 != 58210094820.50`, `58341252840.17 != 58342614440.63`, `58688155764.47 != 58689517364.93`) | la v4 corrige `R1!F328`; la corrida cierra con `O9=0 / P9=TRUE` |
| `MicroFixesHu17Tests.L1_Hu12_BloqueVacio_NoSeLimpia_OutputIgualTemplate` | mismo `ERR-VALIDACION` (el test necesita la corrida Q1 completa) | ídem |
| `MicroFixesHu17Tests.S4_Hu16_InterventoriaLecturaPorAse_Debug_HitoInformation` | ídem | ídem |
| `MicroFixesHu17Tests.W22_Hu14_Validaciones_CulturaEsCo_LineaYEventoCoherentes` | ídem (el rojo era el mismo, con `,` decimal) | ídem |
| `ObservabilidadTests.BloqueValidaciones_PropiedadValidacion_NombreReal_ConAseId` | ídem | ídem |
| `ObservabilidadTests.EjecucionPeriodo_RunIdCorrelacionaInicioAFin_YAseId` | ídem | ídem |
| `ObservabilidadTests.Niveles_HitosInformation_DetalleDebug` | ídem | ídem |

**Causa raíz medida (no inferida).** En la salida de julio-1 con la v4: `Reporte Componentes R1!L213 = 0` y `M213 = 1.361.600,46`. La fórmula de la v3 restaba `M213`; la de la v4 resta `L213`. Diferencia exacta entre las dos plantillas: **1.361.600,46**, el mismo Δ que §12/§13 atribuyeron a `R1!F328`. Es decir: los 7 rojos no eran defectos de la app — eran la plantilla v3.

Tests nuevos (5): `HuellaPlantillaVigenteV4Plan59Tests.V4_1…V4_5` (§14.4), todos verdes.

Renombres (4 pares, mismo cuerpo, mismo veredicto):
`Bases2026072_ConLaV3_…` → `…ConLaPlantillaVigente_…`;
`H1_2_R1_EscribeLosCincoBloquesSobreLaV3_…` → `…SobreLaPlantillaVigente_…`;
`H2_2_…NoTieneFormulaEnLaV3` → `H2_2_…NoTieneFormula`;
`CapaA_Q1_EvaluacionHonesta_RecaudoCierra_EspejoNo` → `CapaA_Q1_EvaluacionHonesta_RecaudoYEspejoCierran`.

**Único aserto que cambió de valor (declarado, no silenciado).** `GoldenValidacionesCruzadasTests.CapaA_Q1…` caracterizaba el defecto E1-2 ("el espejo de remuneración NO cierra": `O25/D34/F34` = FALSE) y con la v4 esos tres sub-bloques evalúan TRUE. No se relajó el gate: el `ValidacionOracleReader` evalúa siempre desde lo escrito y el test sigue exigiendo el resultado verdadero; lo que cambió es la verdad (la plantilla cerró). Se midió el Δ exacto antes de tocar el aserto (§ arriba) y se dejó el motivo escrito en el propio test.

Los 27 rojos restantes son los mismos de la base (mapas LEGACY, fixtures anclados a la v2 por su objeto, paridad CLI histórica): **ni uno cambió de causa**.

### 14.6 Greps y estado del árbol

- **Ruta de la v3 en código productivo:** `0`. La app no conoce plantilla por ruta: la recibe por `--plantilla` / `SolicitudProcesoPeriodo`.
- **Ruta de la v3 en fixtures de procesamiento:** `0`. Solo la abren `GeometriaDerivadaV3T0Plan59Tests` (evidencia congelada del H0, motivo en la cabecera) y `HuellaPlantillaVigenteV4Plan59Tests` (necesita el archivo histórico para probar `v4 = v3 + 2 fórmulas`).
- **Literales de período en producción:** `0` fuera de comentarios/XML-doc y de la exención declarada del texto de ayuda del CLI (§13.1/§13.5).
- **Plantillas byte-intactas:** v4 `5FEDFD31…E38E5` (705.805 B) y v3 `8D0B3299…423538` (709.040 B), idénticas antes y después de todas las corridas.
- **`git status` acotado:** sin cambios en `Docs/**/*.xlsx` (la carpeta `v4/` sigue como archivo nuevo sin trackear, tal como llegó) y sin artefactos `TestResults/` nuevos (los trx de esta tanda se escribieron en `%TEMP%`).

### 14.7 Decisiones que quedan al Ingeniero

1. **Naming del acceso histórico v2 en `Insumos`**: hoy `Insumos.PlantillaOficial()` devuelve la **v2** (histórica) y la vigente es `Insumos.PlantillaOficialV4()`. El nombre del primero quedó engañoso; renombrarlo a `PlantillaHistoricaV2()` toca ~17 fixtures y no cambia comportamiento. No se tocó para no ampliar el diff sin su OK.
2. **Fixtures que quedaron anclados a la v2 por el typo de la plantilla** (§13.4-b: `ProcesadorPeriodoTests` 2026072, `SelladoCacheFormulaT2Tests`, `ParidadCliTests`, `CodigosSalidaCliTests`): el motivo desapareció con la v4. Migrarlos es tentador, pero mueve bases y asertos históricos; se dejan como están y se decide con usted.
3. **`RegresionEvaluaSellaPlan40Tests`** sigue corriendo contra la v2 (`Insumos.PlantillaOficial()`): sus asertos del "defecto espejo declarado" son ciertos **para la v2** y ya no describen la plantilla vigente. Pertenece al mismo paquete de decisión que el punto 2.
4. **Mensaje de producción desactualizado**: la corrida Q2 sigue emitiendo `ASE n: conciliación por empresa omitida en Q2 (residuo R-MAPA-V2-R2-ASE4: la celda destino R2!E285 es fórmula en la plantilla v2; ver Plan 50 §2.4)`. Con la v4 esa celda ya no es fórmula, así que el texto miente aunque el comportamiento sea el vigente. No se tocó (es código productivo y el recorte no es de este frente): requiere decisión sobre si se retira el recorte (Plan 59 §7 lo declaraba habilitado) o solo se corrige el mensaje.

---

## 15. Cierre de los dos últimos frentes (2026-10-27) — conciliación por empresa en Q2 + limpieza de fixtures v2

Ejecutado sobre la **v4 vigente** (`a7476ee2…8c55`), sin commits, sin `checkout`/`reset`, sin escribir en plantillas y sin habilitar `R-SIGMA-Q2` en Q2.

### 15.1 Frente 1 (H3) — retiro del recorte de la conciliación por empresa en Q2

**Premisa verificada por bytes, no heredada.** Con el helper BCL (ZIP+XML) sobre la v4 y la v2:

| Hoja · celda | v2 | v4 |
|---|---|---|
| `Rem. Anticipos R2!E285` | `<f>E192+E203-K192</f>` (fórmula) | **sin `<f>` y sin `<v>`** ⇒ slot de valor puro |
| `Rem. Anticipos R2!K285` | `<f>` compartida | **sin `<f>`** |

El bloqueo `R-MAPA-V2-R2-ASE4` que sostenía el recorte **no existe** en la plantilla vigente, y el 2.º cuadro de R2 declara el grupo `canal A / empresa EAB` en los 5 ASE (§0.4-c/§0.5). Por tanto el recorte era ya solo deuda.

**Diff (2 archivos de producción, 1 de doc):**

| Archivo | Cambio | Traza |
|---|---|---|
| `Remuneracion.Core/Services/ProcesadorPeriodo.cs` | desaparece `if (!esQuincena2) { … } else { mensaje }`: `leaf.Conciliacion = _leafReader.LeerConciliacionEmpresas(...)` corre **igual en ambas quincenas**; comentario de cabecera del método actualizado | §1.2-H3, §2-D-3, S4.1, CA-B3, §0.2.2 |
| `Remuneracion.Core/Services/ValidadorBasico.cs` | **solo comentario**: `R-SIGMA-Q2` sigue cortando en `NumeroQuincena != 1` (código intacto); el texto ya no afirma que el recorte "se mantiene" | §1.3 (R-SIGMA-Q2 **no** se habilita en Q2), §0.2.2 |
| `Remuneracion.Infrastructure/Excel/WorkbookLeafCoherence.cs` | **solo comentario**: mismo gate Σ, misma condición `!= 1`; se documenta que la lista ahora llega poblada en Q2 y que la rama es justamente la que impide reactivarlo allí | ídem |
| `.opencode/project-context.md` | §"Alcance Q2 acotado" → "Conciliación por empresa en Q2 — SIN recorte"; §A2 reescrito; `R-RECORTE-Q2` sale de la lista de residuos vivos | doctrina del proyecto |

**Mensaje retirado:** la cadena `conciliación por empresa omitida en Q2 (residuo R-MAPA-V2-R2-ASE4… )` ya **no existe en producción** (0 ocurrencias en `Core`/`Infrastructure`/`WinForms`/`Cli` fuera de comentarios históricos de planes y docs de evidencia).

**Tests que caracterizaban el defecto (cambiados a exigir lo correcto, con motivo escrito en el cuerpo):**

| Test | Antes | Ahora |
|---|---|---|
| `ProcesadorPeriodoTests.Ejecutar_Periodo2026072_SalidaCertificada5De5_ConDetRetri` | `Assert.Empty(l.Conciliacion)` | exige las **5 empresas del catálogo** por ASE y destinos resueltos en R2. No se fija un conteo exacto: el resolvedor por etiquetas (Plan 55) añade un acumulado sintético por etiqueta real ajena al catálogo — medido: julio-2 ASE3 trae `EAB` (73.523,10) en su R2 ⇒ 6 entradas **para ese ASE** |
| `ConciliacionEmpresaCierreBloqueT0Tests.T0_5_RecorteQ2_NoSeLevanta_DestinoE285EsFormula` | el nombre afirmaba que el recorte "no se levanta" | renombrado `T0_5_PlantillaV2_DestinoE285EsFormula_ClaseRMapaV2R2Ase4`: su objeto son los **bytes de la v2** (evidencia T0.5), no el comportamiento del procesador. Asertos intactos |
| `FuentesAusentesPlan50Tests` (cabecera) | "A2 NO se habilita" | se declara el estado vigente y el párrafo anterior queda marcado como histórico del Plan 50 |

### 15.2 Frente 1 — evidencia de que la conciliación por empresa CORRE en Q2

**a) Log de la corrida real `2026072 --cinco-ase` con la v4** (build FINAL de la tanda; `remuneracion_log_20261027.txt`, runId `12a340ec-737f-46e1-92d5-cdcf46fbdfeb`, 10:56):

```
10:56:25.291 [INF] (RunId=12a340ec… Periodo=2026072 AseId=1) ASE 1: leyendo conciliación por empresa (R1/R2/R4)...
10:56:25.900 [INF] (RunId=12a340ec… Periodo=2026072 AseId=2) ASE 2: leyendo conciliación por empresa (R1/R2/R4)...
10:56:25.958 [INF] (RunId=12a340ec… Periodo=2026072 AseId=3) ASE 3: leyendo conciliación por empresa (R1/R2/R4)...
10:56:26.045 [INF] (RunId=12a340ec… Periodo=2026072 AseId=4) ASE 4: leyendo conciliación por empresa (R1/R2/R4)...
10:56:26.104 [INF] (RunId=12a340ec… Periodo=2026072 AseId=5) ASE 5: leyendo conciliación por empresa (R1/R2/R4)...
```

`5/5` líneas. Las 10 ocurrencias de `omitida en Q2` que quedan en ese log son de las corridas **anteriores** al cambio (08:07 y 08:34, binarios previos); después del cambio hay **0**.

**b) Celdas/hojas (2026072 vs oráculo administrativo `Remuneracion 202607-2 Total Administrativo.xlsx`), comparación celda a celda con tolerancia ±0,5:**

| Hoja | Celdas numéricas comparadas | Diferencias |
|---|---|---|
| `CONSOLIDADO_TOTAL RECAUDO` | 536 | **0** |
| `VALIDACION_TOTAL` | 52 | **0** |
| `Recaudo EAAB Reciprocidad` | 26 | **0** |
| `Recaudo ENEL` | 26 | **0** |
| `Recaudo ENERBIT` | 26 | **0** |
| `Recaudo EAAB + Ciud Limp` | 26 | **0** |
| `Recaudo Directa Occidente` | 26 | **0** |
| `DetRetri2026072` | 91 | 2 — y **no son valores**: `D6`/`D7` son fecha/hora; la app las **sella desde el R10** (Plan 28) y el manual las deja en `0` |

**c) El cruce por empresa en R2 (el bloqueo histórico del `EAB`)**: el grupo `canal A / empresa EAB` queda con valores en el 2.º cuadro de R2 de los 5 ASE. Caso con valor no nulo en julio-2 — ASE3: `R2!E200 = 73.523,10` (cierre del grupo `canal A`) y `R2!E201 = 73.523,10` (`B=EAB / C=Total`), **idéntico al oráculo** (`R2!E161 = 73.523,10`). Es exactamente el monto que el Plan 50 §4 medía como "etiqueta real sin destino en el mapa".

**d) Estabilidad del entregable (la prueba de que el retiro no cambia nada de más):** comparación valor a valor de las **40 hojas** entre la corrida previa al retiro y la posterior, para los tres períodos:

| Período | Hojas con diferencias |
|---|---|
| `2026071` | **0** |
| `2026072` | **0** |
| `2026101` | **0** |

Lectura: el recorte solo impedía **leer** `leaf.Conciliacion`; la escritura del detalle por empresa y de las hojas `Recaudo *` ya estaba cubierta por el detalle derivado R2/R4 (§12) y por A1. Al habilitar la lectura, el cruce por empresa entra además a las validaciones de la fase de cálculo (donde `R-SIGMA-Q2` sigue sin habilitarse en Q2) **sin mover un solo valor del entregable**. No apareció ningún rojo de otra familia: la condición de STOP de §1.4 (R-Q2-RECORTE) no se activó.

### 15.3 Frente 2 (a) — `PlantillaOficial()` pasa a `PlantillaHistoricaV2()`

`Insumos.PlantillaOficial()` devolvía la **v2** mientras la vigente es la v4: nombre engañoso. Renombrado **sin alias** (el nombre viejo ya no existe). **Cero cambio de comportamiento** (misma ruta, mismo archivo).

- Consumidores actualizados: **16 fixtures** (29 usos) + `Insumos.cs`: `ComposicionServiciosTests`, `ConciliacionEmpresaSigmaPlan55Tests`, `ContratoSeccionPorOperandoR1Tests`, `ControlAntiFalsoCeroPlan50Tests`, `DivergenciaDeclaradaCierreT4Tests`, `DivergenciaDeclaradaCliTests`, `DivergenciaDeclaradaReejecucionUiTests`, `DivergenciaDeclaradaT0Tests`, `FNew1Q1BlockerT0Tests`, `FuentesAusentesPlan50Tests`, `HuellaAceptacionO3Tests`, `IntegridadPaquetePlan58Tests`, `RegresionOficialPlan39Tests`, `ReporteBancoFuentesOpcionalesPlan51Tests`, `SalidaSinValidarO4Tests`, `SeccionPorOperandoR1T0Tests`.
- Helpers locales que también mentían: `FNew1Q1BlockerT0Tests` pasa a llamar `Insumos.PlantillaHistoricaV2()` directo (se elimina el wrapper); `TablaJS26S31T3Tests` apuntaba a la **v1** ⇒ `PlantillaHistoricaV1()`; `AnclasDerivadasH2Plan57Tests` apuntaba a la **v4** ⇒ `PlantillaVigente()`.
- `Insumos.BaseDelPeriodo` (base v2 vía Phase 0) conserva su nombre y su firma —lo usan `BaseV2DelPeriodo` y las caracterizaciones de geometría v2— con la documentación corregida para no llamarla "oficial".

### 15.4 Frente 2 (b) — fixtures migrados de v2 a v4 (motivo del anclaje: el typo `F328`)

Cada expectativa se conserva porque **la causa del anclaje era de plantilla, no de la app**: la v4 es la v3 con `R1!F328` (`-M213` → `-L213`) y `R1!F333` (`--L126` → `-L126`) corregidas (§14.4), misma geometría, mismas etiquetas, mismos valores. Los tres períodos cierran con la v4 (§14.3 y §15.5 de este cierre).

| Fixture | Antes (base) | Después (base) | Por qué la expectativa se conserva / qué cambió |
|---|---|---|---|
| `ProcesadorPeriodoTests.Ejecutar_Periodo2026072_ConLectorOracle_VeredictosValidacionesNoVacios` | `BaseV2DelPeriodo("2026072")` | `BaseDelPeriodoV4("2026072")` | Exige `5` bloques `VALIDACIONES:` (sin afirmar `TRUE`): con la v4 la corrida **cierra** y los 5 veredictos siguen presentes. Respaldo: oráculo administrativo de julio-2 + corrida real exit 0 |
| `SelladoCacheFormulaT2Tests` (flujo Q2 sano y sintético) | `BaseV2DelPeriodo("2026072")` | `BaseDelPeriodoV4("2026072")` | Exige `C15/D25/O25/D34/F34 = "1"` en el flujo sano y `C15="0"` en el sintético (Δ=1 inyectado), con `sellado == evaluación`. Con la v2 el typo de `F328` habría sellado `0`: con la v4 sella el valor verdadero. Respaldo: oráculo de julio-2 |
| `ParidadCliTests.Paridad_5Ase_Q1` y `…_Q2` | `BaseV2DelPeriodo("2026071"/"2026072")` | `BaseDelPeriodoV4(…)` | Exige `CodigosSalida.Ok` en CLI y vía directa. **Además se sustituyó el mecanismo de comparación**: recorría `WorkbookLeafCellMapPorAse`/`WorkbookLeafCellMapQ2` + `RolsQ2` (foto de la v2), y con la v4 esas direcciones ya no describen los roles ⇒ el rojo era del fixture. Ahora compara **libro contra libro, celda por celda, en las 40 hojas** (`AssertMismoContenidoPorValor`): más fuerte que la lista congelada, cero coordenadas. Respaldo: R10 de julio-1 y administrativo de julio-2 |
| `CodigosSalidaCliTests.Ejecucion_5Ase_Q1_Devuelve0_ConArchivoYResultadoOk` | `BaseV2DelPeriodo("2026071")` | `BaseDelPeriodoV4("2026071")` | Exige exit 0 + `RESULTADO OK codigo=0` + archivo + `runId=`; nada se relajó. Era el rojo por el Δ 22.921,61 de la v2. Respaldo: R10 de julio-1 |
| `RegresionEvaluaSellaPlan40Tests` | `PlantillaHistoricaV2()`, con ramas Q1 (`ERR-SIN-SLOT` + espejo en FALSE) y Q2 (divergencia sintética) | `PlantillaOficialV4()`; **renombrado** `T4_EndToEnd_PlantillaVigente_EvaluaSella` | El "defecto espejo declarado" (E1-2) ya **no existe**: el caso Q1 se reescribió para exigir el comportamiento correcto (corrida que cierra, nombre definitivo y los 5 sub-bloques en `1`, con `sellado == evaluación` leído con `ValidacionOracleReader`). El caso Q2 conserva su objeto: el gate se detiene ante la divergencia **sintética** (`SINTETICA-C15`, Δ = 1) y el sellado se prueba en ambos sentidos. Respaldo: R10 de julio-1 + administrativo de julio-2 |

**Ningún test se borró, se relajó ni se marcó `Skip`.** La única baja de la lista de rojos es un cierre real (`CodigosSalidaCli…`, y los dos `ParidadCli*`).

### 15.5 Verificación exigida

- **Build**: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- **Corridas reales con la v4** (`--cinco-ase`, salida en ruta temp FRESCA `%TEMP%\p59-f1-run2\<período>`; la tanda previa de la misma evidencia quedó en `%TEMP%\p59-f1-run\<período>`):

| Período | Exit | Veredicto | `DetRetri!D9:D14` |
|---|---|---|---|
| `2026071` | **0** | `RESULTADO OK` | 16.758.167.585 / 12.227.292.053 / 10.127.195.599 / 10.553.776.376 / 8.543.663.209 / **58.210.094.822** |
| `2026072` | **0** | `RESULTADO OK` | 17.450.228.673 / 20.516.143.970 / 15.221.896.467 / 7.179.595.396 / 12.073.344.662 / **72.441.209.168** |
| `2026101` | **0** | `RESULTADO OK` | 14.842.425.136 / 11.302.637.605 / 8.823.787.300 / 10.483.172.762 / 7.550.668.665 / **53.002.691.468** |

  Los tres: `VALIDACION_TOTAL O9 = 0; P9 = TRUE` en los 5 ASE, 0 `ERR-SIN-SLOT`, 0 `ERR-PLANTILLA`, 0 `ERR-ESCRITURA`, 0 `ERR-INTEGRIDAD`. Idénticos a §14.3 ⇒ **el retiro del recorte no movió un peso del entregable**.

- **Suite completa y set-diff** (trx `%TEMP%\p59-f1-final2\p59-f1.trx`; base de la tanda, medida en el mismo árbol ANTES de los cambios: `%TEMP%\p59-f1-base\p59-base.trx`):

| | Total | Verdes | Rojos | Skip |
|---|---|---|---|---|
| Base medida de esta tanda | 1623 | 1596 | 27 | 0 |
| **Con los dos frentes cerrados** | **1623** | **1599** | **24** | 0 |

  **Set-diff: 0 rojos nuevos · 3 rojos cerrados · 0 tests nuevos · 2 renombres.**

  Cerrados (una sola causa, medida): `CodigosSalidaCliTests.Ejecucion_5Ase_Q1_Devuelve0_ConArchivoYResultadoOk` (el Δ 22.921,61 de la v2 en `VALIDACION_TOTAL`), `ParidadCliTests.Paridad_5Ase_Q1_CliVsDirecto_ValoresYSnapshotsIguales` y `ParidadCliTests.Paridad_5Ase_Q2_CliVsDirecto_IncluyeAjustesYDetRetri` (rojo del mecanismo de comparación congelado + base v2). Los **24 restantes son los mismos de la base y con la misma causa** (mapas LEGACY, caracterizaciones cuya base es la v2 a propósito, fixtures de agosto/permuta BCE): ninguna tolerancia ampliada, ningún gate retirado, ningún test borrado, ningún `Skip`.

  Renombres (mismo cuerpo y mismo veredicto, nombre honesto): `T4_EndToEnd_Oficial_EvaluaSella_DefectoEspejoDeclarado` → `T4_EndToEnd_PlantillaVigente_EvaluaSella`; `T0_5_RecorteQ2_NoSeLevanta_DestinoE285EsFormula` → `T0_5_PlantillaV2_DestinoE285EsFormula_ClaseRMapaV2R2Ase4`.

- **Greps de cierre**:
  - `0` ramas por quincena en el bloque tocado (`ProcesadorPeriodo`): no queda ningún `if (esQuincena2)` alrededor de la conciliación; los `esQuincena2` restantes son de **dominio de negocio** (AJUSTES-SF-T solo en 2.ª quincena, `CalcularConsolidado`, ajustes por ASE) y estaban fuera de alcance.
  - `0` referencias en producción al mensaje retirado y `0` menciones al recorte como vigente (solo quedan comentarios de `R-SIGMA-Q2`, que **no** se toca).
  - `0` literales de período en producción fuera de la exención declarada del texto de ayuda del CLI (`OpcionesCli.cs:219`), verificado con el escaneo del candado (`CandadoPeriodoQuemadoPlan59Tests`, verde en la suite).
  - `0` ocurrencias de `Insumos.PlantillaOficial(` (el nombre viejo no existe).
  - **Plantillas byte-intactas**: v4 `5FEDFD31…0A4E38E5` (705.805 B), v3 `8D0B3299…2B423538` (709.040 B), v2 `86E48E47…CC9D3D14` (693.983 B), v1 `3086AC31…A49C5214` (692.598 B).
  - `git status`: sin cambios en `Docs/**/*.xlsx` (la carpeta `v4/` sigue como archivo nuevo sin trackear, tal como llegó) y sin trx nuevos en `TestResults/` — los tres `.trx` sin trackear que se ven allí (`p59-frente3-final.trx`, `p59-h5-cierre.trx`, `p59-h5-suite.trx`) son de la tanda **anterior** (26-oct 21:09 y 27-oct 08:06/08:16); los de esta tanda se escribieron en `%TEMP%` (`p59-f1-base`, `p59-f1-final`, `p59-f1-final2`).
- **CRLF y codificación**: los archivos tocados quedan con CRLF uniforme y **0 caracteres `U+FFFD`** (verificado leyendo cada archivo como UTF-8 y contando el reemplazo). Se normalizaron a CRLF cinco archivos que la tanda anterior había dejado mixtos o en LF (`Docs/T0-Plan59-H2bis-Detalle-Derivado.md`, `DetallePeriodoDerivado.cs`, `WorkbookLeafCellMapBalanceSc.cs`, `DetalleR2R4EscrituraTests.cs`, `GoldenDetRetriQ2Tests.cs`) **sin cambiar su contenido** (mismo texto, solo fin de línea).

### 15.6 Puntos que quedan al Ingeniero

1. **`R-SIGMA-Q2` sigue vivo y sin habilitar en Q2** (§1.3). Ahora que la lista de conciliación llega **poblada** en Q2, el gate es *tentador*: medido, la descomposición sigue sin ser válida allí (R4 no declara todas las empresas, R2 trae etiquetas reales sin destino en el catálogo). Reactivarlo exige contrato nuevo + T0 propio.
2. **`LeerRecaudoEmpresa` elige el par de columnas del `RESUMEN MES` por quincena** (D/E vs F/G, `ExcelDataReaderWorkbookLeafInputReader.cs:1110-1123`). No se tocó: es el **layout de la fuente** (dos pares de columnas rotulados `VALOR 1°Q`/`VALOR 2°Q`), no una regla de negocio, y no participa del retiro del recorte. Queda anotado por si el Ingeniero quiere que también eso se derive por rótulo de encabezado en un frente propio.
3. **`WorkbookLeafCellMapQ2` / `RolsQ2` / `WorkbookLeafCellMapPorAse` ya solo los consumen fixtures LEGACY.** `ParidadCliTests` dejó de usarlos en este cierre; quedan `MapaQ2T0Tests`, `DetalleR2R4EscrituraTests` y `GoldenDetRetriQ2Tests`, cuya base es la v2 **a propósito** (su objeto ES la foto). Recomiendo que el siguiente frente los marque `[Obsolete]` con motivo o los retire cuando su evidencia histórica deje de citarse.
4. **`RegresionEvaluaSellaPlan40Tests` cambió de nombre y de veredicto**: pasó de "declara el defecto espejo" a exigir el cierre. Si algún documento de trazabilidad (matriz de requisitos del Plan 40) cita el nombre viejo `T4_EndToEnd_Oficial_EvaluaSella_DefectoEspejoDeclarado`, hay que actualizarlo: el defecto E1-2 quedó cerrado por la plantilla, no por el código.
