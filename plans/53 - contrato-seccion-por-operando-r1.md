# Plan 53 — Contrato de sección por operando/celda R1 (fin del ordinal, rediseño Ruta B)

| Campo | Valor |
|---|---|
| Estado | **CERRADO 2026-10-22** (H0-H3 verdes; code-review APROBADO CON OBSERVACIONES; 2 desviaciones ratificadas — ver §13) |
| Fecha | 2026-10-22 |
| Base de regresión | `p51-h2-suite.trx` (**1260 / 1193 / 67 / 0**) + `p51-h2-setdiff.txt` |
| Familia | Opcionalidad de fuentes/canales/secciones (reglas 6 y 7); hijo de rescate del Plan 52 (CERRADO CON STOP §13) |
| Toca | `WorkbookLeafCellMapPorEmpresa.cs:28-128` (contrato `EditablesR1PorEmpresa` + enum `FuenteR1` o nuevo `SeccionEsperada`) + `ExcelDataReaderWorkbookLeafInputReader.cs:1560-1640` (`ExtraerCeldasR1`) + `ContratoSeccionPeriodFreeR1.cs` (solo consumo, sin cambiar `SeccionPorCierre`/`Alinear`) |
| No toca | Plantillas v1/v2 (SHA §8), writer (`OpenXmlPlantillaWriter`), banco (`ReporteBanco*`), mapas R2/R4 (`EditablesR2PorEmpresa` / `EditablesR4PorEmpresa`), `CalcularVisibleR1` (lógica), Planes 50/51/52 y sus deudas (`R-SIGMA-Q2`, `R-MAPA-V2-R2-ASE4`, `R-PARCIAL-OCT-ASE3-R1-OCC` origen) |

---

## 0. Clarification Gate (pre-respondido por el encargo — se registra, no se re-pregunta)

El encargo del Ingeniero ya arbitra lo que en el Plan 52 era bloqueo; se deja por escrito para que el §12 solo pida lo que falta:

1. **Alcance cerrado A1–A4** (§3): solo operandos `EditablesR1PorEmpresa`; R2/R4 solo se declaran (D6), no se tocan. Banco, writer, plantillas: intactos.
2. **T0 read-only bloqueante** (§10-T0 / §4-T0): el contrato celda → sección esperada se **deriva de la plantilla/manual**, no se inventa; se cruza con la matriz de 180 filas-empresa del T0 del Plan 52 y con los 4 períodos (jul-1, jul-2, ago-2, oct-1). Si una celda no se puede tipar con evidencia → **STOP y reporte**.
3. **Reutilizar, no inventar** (§5-D3): predicado vigente `SeccionPorCierre` (Plan 46); `PorFronteraArriba` (Planes 41/45) solo como contraste de partición/orden. Cero lógica nueva de sección.
4. **Tests no se debilitan**: 0 `Skip`, 0 `Assert.True(true)`, 0 borrados. La lista cerrada de Δ se congela en el T0 (los 17 del T0-52 + los que aparezcan); **rojo fuera de lista = STOP**.
5. **Mapeo por ETIQUETA/SECCIÓN**, nunca por posición/ocurrencia (reglas 6 y 7). `OrderByDescending(|F|)` sale de este camino.

---

## FASE 1 — PROPOSE

### 1. Contexto y doctrina aplicable (no se re-discute)

- **Reglas 6 y 7 del project-context** (reunión con la propietaria 2026-10-21/22): las fuentes **Y** los canales pueden venir o no venir; una sección/canal ausente **no es estructura rota**; su celda destino queda en **0**. Prohibido resolver por ocurrencia/índice. **La señal es la frontera** (`Mes`, `Aplicacion`, `Componente`), nunca el índice. Este plan extiende el mismo método al eje sección-operando (Plan 51: precedente de header-scan en banco).
- **Plan 52 §13 (CERRADO CON STOP 2026-10-22)**: el intento anterior (resolver por ordinal mayor/menor/intermedio sobre secciones presentes) es correcto como *dirección* pero insuficiente como *contrato*: la sección no se puede inferir por ordinal ni por magnitud. Camino ordinal **prohibido a partir de este plan**. La heurística `TotalMain=[0] / TotalMenor=[count-1] / TotalMedio=[1]` sobre listas ordenadas por `|F|` (`ExcelDataReaderWorkbookLeafInputReader.cs:1585-1637`) queda marcada para retiro en este camino.
- **Planes 41/45/46** (`SeccionEspejoR1` / `ContratoSeccionPeriodFreeR1`): contrato de sección vigente y reutilizable. **Plan 46 supersede al 45** en tipado (cierre + canal efectivo); Plan 41 §E1-1/E1-3 (prohibición de fallback posicional, `ERR-SIN-SLOT`) sigue vigente.
- **Plan 50 §8, residuo 1** (`R-PARCIAL-OCT-ASE3-R1-OCC`, reclasificado 2026-10-22): NO es estructura rota — la fuente OCCIDENTE no trae la primera sección (`AFaseo`) en octubre-1 y eso es normal. El defecto es el **mapa por ocurrencia**. Fix en plan propio (= este plan). Plan 51 (misma familia de opcionalidad, otro reporte): precedente de método — resolver por **etiqueta observada**, ausente = 0 con huella, sin ramas por período/empresa.
- **T0 del Plan 52** (`Docs/T0-Plan52-Seccion-Operandos-Empresa-R1.md`, 299 líneas): **fuente única de verdad del diagnóstico**. 52 grupos, 180 filas-empresa, 4 períodos × 5 ASE, tipabilidad total, identidad, manual. Este plan **no re-investiga**: adopta su matriz (§2), su S1.5/S1.6/S1.7 y sus 17 Δ (§6), y solo añade lo que el encargo ordena: el **contrato celda → sección esperada**.

### 2. Problema (probado; no re-investigar)

La resolución actual de los operandos por empresa usa **índices de ocurrencia** sobre listas ordenadas por `|F|` descendente (`ExcelDataReaderWorkbookLeafInputReader.cs:1585-1637`):

| Operando | Resolución actual | Supuesto oculto |
|---|---|---|
| `TotalMain` / `EspecialesMain` | `filasTotal[0]` | el mayor `\|F\|` es el bloque oportuno |
| `TotalMenor` / `SubsMenor` / `EspecialesMenor` | `filasTotal[count-1]` / `filasSubs[count-1]` | el menor `\|F\|` es la primera sección |
| `TotalMedio` | `filasTotal[1]` | siempre hay ≥ 2 filas |

El T0 del Plan 52 probó que **la sección no se puede inferir por ordinal ni por magnitud** (T0 §0–§6, veredicto ROJO):

- **Oct-1 ASE5**: ENEL `7.414.210.944,89`, ENERBIT `160.186.886,97`, OCCIDENTE `134.612.117,14`, cada una en **una sola fila de la primera sección** (la única de recaudo de la hoja, `Oportuno#1`). `Main=[0]` con regla D3-del-52 (`única primera ⇒ Menor, Main=0`) **borra 160 millones** (ENERBIT no declara `TotalMenor` en el mapa `(3,5)`) y deja `VisibleR1` negativo (`−20.676.306,97`). Ninguna decisión ratificada cubría este caso (T0 §9.1-3).
- **Oct-1 ASE3 OCCIDENTE**: solo el bloque oportuno `286.711.907,53` (`Oportuno#2`, R23); la primera sección (`AFaseo`) **no vino** y es normal → `Menor = 0`, no `count-1` (que duplicaría el bloque grande en `F227 = F253+F237+F216−L237`).
- **Ago-2 ASE2 ENEL/OCCIDENTE**: 2–3 filas con **magnitud ≠ ordinal** (`−1.414.859,14` vs `131.543,19`; `−28.650,84` vs `24.890,00`). Verdes hoy, el plan-52 los movía sin regla aprobada → STOP por diseño (T0 §3).
- **Jul-2 ASE3**: el encargo suponía 1 fila en primera sección; el T0 midió **2** (`f7 Oportuno#1 = −16.458,46` y `f28 Oportuno#2 = 825.593.716,53`) — el test D5-2 del Plan 52 no aplica; el Δ real es solo `TotalMedio → 0`.
- El manual confirma la resolución por sección en 4 casos (jul-2 ASE5 ENEL, ago-2 ASE5 ENEL, ago-2 ASE2 ENEL, ago-2 ASE2 OCCIDENTE): **el código actual es el defecto** (duplica el bloque), la regla por sección corrige al centavo (T0 §8-a).
- `SeccionPorCierre` vs `PorFronteraArriba` difieren en etiqueta en 180/180 pero **coinciden en partición y orden** (T0 §7): la resolución solo necesita (partición, orden, "primera vs no primera"). La premisa D1-del-52 ("cierre == apertura por construcción") es falsa y debe reescribirse citando Plan 46 §2.2 (DEFECTO 2).

### 3. Alcance APROBADO (Ruta B — rediseño, no parche; prohibido ampliar)

- **A1. Contrato explícito de sección por operando/celda.** Enriquecer `EditablesR1PorEmpresa`: cada entrada declara **qué sección espera** (rol de sección) además del tipo de valor. El DESIGN (§5-D1) evalúa (a) `(Celda, FuenteR1, SeccionEsperada)` con enum nuevo, (b) `RolOperando` que combina tipo + sección, (c) mapa paralelo `Celda → Sección`, y recomienda. El contrato se **deriva de la plantilla/manual**, no se inventa: la celda destino (`F237`, `F216`, `F222`, `F487`, etc.) tiene una posición que corresponde a un bloque/sección. El T0 lo congela celda por celda.
- **A2. Resolución por sección real, no por ordinal.** La sección real de cada fila fuente se obtiene con el predicado vigente (`SeccionPorCierre`, Plan 46). `Main`/`Menor`/`Medio`/`Subs`/`Especiales` se resuelven por su **sección declarada**, no por `[0]`/`[count-1]`/`[1]`. **Si la sección declarada no existe en la fuente → 0** (la fuente no la trae; es normal). **Si existe más de una fila en la misma sección → STOP** (ambigüedad, no heurística). **Si la sección declarada es la única de recaudo (ASE5 oct) → alimenta `Main`**; el contrato lo refleja celda por celda (el mapa `(3,5)` sin `TotalMenor` es la prueba de fuego).
- **A3. Invariantes intactos.** `A='Componente' B='Total'`, `A='Subs/Cont' B='Total'`, `A='Total'` siguen fail-fast. La ausencia de fuente/canal/sección **no** es fail-fast.
- **A4. No tocar plantillas ni writer.** v2 SHA `0F192AD2721C77986B6A8CB85D1A92F782DCD2F699F362E066FA17027F2DAB`; v1 SHA `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`. Values-only.

### 4. No-objetivos (STOP si aparecen)

1. No tocar los Planes 50/51/52 ni sus deudas (`R-SIGMA-Q2`, `R-MAPA-V2-R2-ASE4`, F-NEW-1).
2. No tocar el banco (`ReporteBanco*`), ni los mapas `EditablesR2PorEmpresa` / `EditablesR4PorEmpresa` (D6 solo declara con T0 propio).
3. No renombrar/reordenar el enum `FuenteR1` de forma rompedora (lo consumen mapa, reader-switch, `CalcularVisibleR1` y tests). El cambio de contrato es **aditivo** (nuevo campo/enum), nunca reordenamiento silencioso.
4. No cambiar columnas de lectura (`F`=col 5, `Especiales`=col 11 fijas) ni el header-scan de componentes: no son el defecto; son frente de otro plan.
5. No reintroducir fallback posicional/ordinal (Plan 41 E1-1) bajo ninguna forma — ni `[count-1]`, ni `[0]`, ni `OrderByDescending(|F|)` en este camino.

---

## FASE 2 — DESIGN

### D1. Modelo del contrato: (a) vs (b) vs (c) — decisión

| Candidata | Forma | Veredicto |
|---|---|---|
| **(a) Tupla con enum de sección** `(string Celda, FuenteR1 Fuente, SeccionEsperada Seccion)` | Un solo mapa; el compilador exige la sección por entrada | **RECOMENDADA** |
| (b) `RolOperando` que combina tipo + sección | Un solo miembro por (tipo, sección), p. ej. `TotalMainUltimaOportuna` | Descartada como principal |
| (c) Mapa paralelo `Celda → Sección` | Segundo diccionario `(empresa, ase, celda) → SeccionEsperada` | Descartada |

**Razonamiento:**

1. **(a) es la única que hace imposible el olvido de ASE5-oct por construcción.** Cada celda declara su sección esperada en el mismo literal que la crea; el review ve `("F487", TotalMain, UnicaRecaudo)` junto al comentario de plantilla/manual. (c) permite que el mapa principal y el paralelo se desincronicen (R-ATRIBUCION: una celda sin entrada en el paralelo caería a un default silencioso o a un throw tardío). (b) multiplica el enum por el producto cartesiano (tipo × sección ≈ 9×3) y obliga a renombrar miembros que consumen `CalcularVisibleR1` y tests — ruptura de contrato sin ganancia.
2. **OCP/D-H:** (a) añade un campo con default documentado y una tabla de roles **relativos** (no literales de período/empresa en producción). Una forma futura con sección nueva = fila citada en T0 nuevo, cero ramas. (b) exigiría versionar el enum por cada forma; (c) exigiría mantener dos mapas sincronizados por cada forma.
3. **Trazabilidad celda → evidencia:** (a) permite el comentario por línea `// F487 ← Oportuno única-sección (oct-1 ASE5 f12; manual …; T0 §…)` en el mismo sitio que el `grep` de cierre audita. (c) dispersa la evidencia en dos archivos.

**Forma normativa de (a) (a congelar por el T0, sin literales de período en producción):**

```csharp
public enum SeccionEsperadaR1
{
    /// <summary>La única sección de recaudo de la hoja (primera == última; caso ASE5-oct).</summary>
    UnicaRecaudo,
    /// <summary>Primera sección presente del bloque (Totales o Subs, según el operando).</summary>
    Primera,
    /// <summary>Última sección oportuna presente del bloque.</summary>
    UltimaOportuna,
    /// <summary>Sección intermedia (solo si hay ≥ 3 secciones presentes).</summary>
    Intermedia,
}
```

- El reader agrupa las filas-empresa del bloque (Totales con `indice < indiceComponente`, Subs con `indice > indiceComponente`, misma regla de firma `C=label ∧ D=Total` vigente) por `SeccionPorCierre`, en orden de aparición. Cada operando toma su fila de su sección declarada; ausente → `0m`; 2 filas en la misma sección → **STOP** (`CalculoInvalidoException` que nombra ASE + empresa + sección + filas, nunca elección silenciosa).
- `Block5Total` (canal `D=5`) conserva semántica intacta (sin sección).
- Si el T0 demuestra que un rol relativo no basta para alguna celda (p. ej. una celda que exige `Aplicacion#k` absoluta), el T0 lo declara como **STOP de modelado** (no se inventa un quinto rol sin evidencia).

### D2. Cómo se deriva la sección esperada de la plantilla (evidencia obligatoria)

**Decisión: por la geometría del bloque de la celda destino + la fórmula que la consume + el manual/oráculo, en ese orden, con cita triple por celda.**

1. **Geometría (primaria):** la fila de la celda destino dentro del bloque del ASE (`RangoDatosAse`: ASE1 6-72, ASE2 110-210, ASE3 289-355, ASE4 450-550, ASE5 603-676) determina su bloque/sección oficial. Ejemplo: `F237` (TotalMain OCC ASE3) vs `F216` (TotalMenor) vs `F222` (TotalMedio) — el T0 lee por BCL (ZIP+XML) la sección oficial que contiene cada celda y la frontera que la cierra.
2. **Fórmula consumidora (corroboración):** la `<f>` que consume la celda (p. ej. `F336 = F253+F237+F216−L237` para OCC ASE3; `F529 = F502+F483+F474−L483` para ENEL ASE5) confirma el rol: el sumando mayor/oportuno vs el sumando primero. El T0 cita el texto-`<f>` exacto (nunca `<v>`).
3. **Manual/oráculo (árbitro):** `REMUNERACION_<empresa>!D{8+ase}` (jul-2/ago-2) y `CONSOLIDADO_TOTAL RECAUDO!D47:D51` deciden los casos donde la geometría es ambigua. Los 4 pares arbitrados a favor de la sección (T0-52 §8-a) son el oráculo de que el defecto es el código actual.

Si una celda no se puede tipar con esta triple evidencia → **STOP y reporte** (encargo T0). El T0 publica la tabla celda → sección esperada para las 13 claves del mapa (conteo: ASE1 3+6, ASE2 3+4+4, ASE3 4+6, ASE4 3+6+6, ASE5 4+3+4 = 56 entradas) con fila-fuente + sección-`SeccionPorCierre` + celda-slot + sección-contrato + canal.

### D3. Qué predicado de sección usar y cómo reconciliar las etiquetas divergentes

**Decisión: `ContratoSeccionPeriodFreeR1.SeccionPorCierre` como tipador único; `SeccionEspejoR1.PorFronteraArriba` solo como contraste de partición/orden.**

Razonamiento (corrige la premisa D1-del-52 citando Plan 46 §2.2 = DEFECTO 2):

1. Las filas-empresa (`C=label ∧ D=Total`, `indice ⋚ indiceComponente`) son **filas de datos, nunca fronteras**: su firma es disjunta de `EsMesTotal`/`EsAplicacionTotal`. `SeccionPorCierre` las tipa por la frontera que las **cierra** (la inmediatamente debajo), lo que distingue una "primera" que cierra en `Mes` (`Oportuno#1`) de una que cierra en `Aplicacion` (`Aplicacion#1`) — distinción que el ordinal no puede hacer.
2. En 180/180 filas-empresa `SeccionPorCierre` y `PorFronteraArriba` devuelven **etiquetas distintas** (la sección se abre por una frontera y se cierra por otra) pero inducen **exactamente la misma partición y el mismo orden** (T0-52 §7). La resolución de operandos solo necesita (partición, orden, "primera vs no primera"), luego el desacuerdo de etiquetas **no altera ningún operando**. El T0 de este plan lo re-verifica fila a fila (S1.1) y cualquier desacuerdo de partición/orden es STOP.
3. `CanalEfectivo` devuelve `null` para estas filas (traen `D=Total` propio) — documentado como degradación esperada: aquí la empresa ya viene discriminada por la **firma** (`C=label`), así que la llave efectiva es **(firma-empresa + sección-por-cierre)**. No se crea variante nueva del predicado. `Alinear` (monotónico fuente→oficial) **no se usa** en este camino (alinea secciones del espejo interior, no bloques-empresa del lector de conciliación); se cita como referencia, no se invoca.
4. El reader trabaja hoy sobre `List<object?[]>` crudas y `SeccionPorCierre` exige `IReadOnlyList<FilaEspejoR1>`: el SPEC (§7-H1.2) reutiliza la **construcción existente** de `FilaEspejoR1` (la que alimenta `LeerEspejoR1`), prohibido reimplementar el scan de fronteras. D-H intacto: el predicado no recibe período/empresa.

### D4. `TotalMedio`: ¿desaparece, pasa a 0, o se mantiene con sección propia?

**Decisión recomendada: se mantiene como operando con sección propia `Intermedia` (≥ 3 secciones presentes); con < 3 secciones vale 0. Los 2 Δ (`F222` jul-2/ago-2 ASE3-OCCIDENTE: `+16.458,46` y `+522,80`) quedan como divergencia declarada con dueño (ratificación §12-1 del Plan 52, extendida).**

Justificación con el manual (T0-52 §8-b):

- `TotalMedio` **no** participa de `CalcularVisibleR1` en ningún ASE (la cadena solo usa `TotalMain`, `EspecialesMain`, `SubsMain`, `TotalMenor` y `Block5Total`), pero **sí** se escribe como literal en `F222`/`F362`/`F368`.
- En la v2 esas celdas son **hojas inertes**: ningún `<f>` las consume (`F29`, `F242`, `F68`, `F338` sin fórmula) y `CONSOLIDADO!D47:D51` apunta a otra cadena (`F77/F256/F417/F570/F682`, p. ej. `F417 = F342+F315−L315`) que el mapa por empresa **no** alimenta — de ahí que `CONSOLIDADO!D47:D51 = 0` en las salidas actuales (deuda abierta del Plan 49, no de este plan).
- El manual (`D47:D51` jul-2 = 49.580 / 56.190.070 / 1.307.680 / 4.343.101 / 441.849) **no** se reproduce con ninguna de las dos hipótesis (`Medio=0` vs `Medio=Menor`) porque su geometría de anclaje es la legacy (R-TRAMPA-ANCLA). Luego `Medio=0` queda **sin contradicción** del manual pero **sin confirmación**: aplica la decisión ratificada del Plan 52 §12-1 (divergencia declarada con dueño).
- Alternativa descartada (eliminar `TotalMedio` del mapa): rompería el contrato celda-a-celda con la plantilla (la celda existe y hoy recibe un literal) y exigiría re-congelar la geometría legacy; el costo supera al beneficio. El T0 congela los 2 Δ exactos; cualquier otro Δ de `Medio` fuera de lista = STOP.

### D5. Lista cerrada de Δ y tratamiento de los 4 casos del manual

**Decisión: la lista cerrada es exactamente la tabla de 17 Δ del T0-52 §6 (reproducida en §9), con el arbitraje del manual §8 como evidencia de que el plan corrige. Los 4 casos del manual se tratan por sección (el código actual es el defecto):**

| Caso del manual | Hoy (ocurrencia) | Plan 53 (sección declarada) | Veredicto |
|---|---|---|---|
| Jul-2 ASE5 ENEL | 24.524.354.313,16 | 11.701.157.844,90 | **Plan = manual** (duplicado `TotalMenor` por `count-1` eliminado; `F474 → 0`) |
| Ago-2 ASE5 ENEL | 25.008.738.395,88 | 11.748.395.433,49 | **Plan = manual** (ídem) |
| Ago-2 ASE2 ENEL | 19.239.204.490,16 | 19.237.658.087,83 | **Plan = manual** (`TotalMenor` por posición f7, no por magnitud f17) |
| Ago-2 ASE2 OCCIDENTE | 2.217.333.974,79 | 2.217.280.433,95 | **Plan = manual** (ídem, f12 vs f20) |

Casos obligatorios del encargo (tratamiento normativo):

- **Oct-1 ASE5 (primera = única = principal):** las 3 empresas traen una sola fila en `Oportuno#1` (única sección de recaudo). El contrato declara sus celdas `Main` con `SeccionEsperada = UnicaRecaudo` → **`Main` = valor, `Menor`/`Medio` = 0 donde el mapa los declare, sin pérdida**. En particular `(3,5)` sin `TotalMenor`: `F487 = 160.186.886,97` intacto. La regla D3-del-52 (`única primera ⇒ Menor`) queda **derogada** y sustituida por el contrato por celda.
- **Oct-1 ASE3 OCCIDENTE (solo oportuna):** `f23 Oportuno#2 = 286.711.907,53`; la primera sección no vino → `TotalMain = valor`, `TotalMenor = 0`, `TotalMedio = 0`, `SubsMenor = 0`. El throw `S2_1` vigente se **reescribe** a aserciones de valor (no se borra el método).
- **Ago-2 ASE2 (2–3 filas, magnitud ≠ ordinal):** `TotalMenor`/`TotalMedio` por **posición de sección** (f7/f17), no por magnitud. Son verdes hoy y cambian por corrección probada por el manual → entran en la lista cerrada (§9 #6–#7), no son regresión.
- **Jul-2 ASE3 (el encargo suponía 1 fila primera):** el T0 midió **2 filas** → no hay caso "1 fila primera" aquí; el Δ es solo `TotalMedio → 0` (§9 #3). Si el T0 de este plan encontrara una 1-fila-primera real en otra combinación, aplica la regla general (sección declarada; ausente → 0).
- **ASE1/ASE4:** el T0-52 no midió divergencias fuera de `SubsMenor` con 1 sola sección de Subs (§9 #1, #5: jul-1/ago-2 ASE1-OCCIDENTE `F29`, ago-2 ASE3-OCCIDENTE `F242`). Esos Δ son **solo de literal escrito** (hojas inertes en v2, §8-c) pero **sí** entran en la lista cerrada porque alteran el diff values-only.

### D6. R2/R4: ¿el mismo defecto existe?

**Decisión: se declara, no se amplía. Veredicto adoptado del T0-52 §5 (S1.4):**

- **R2 no tiene caso real en los 4 períodos**: 50/52 pares con exactamente 2 filas (una en `Componente`, una en `Subs/Cont`); 2 pares con 0 filas (ENERBIT ASE5 jul-1 y oct-1, ausencia total Plan 50 A1). **Ningún par con 1 fila.** Luego el defecto gemelo (`Total=[0]` y `SubsCont=[count-1]` colapsando) **no** tiene caso real. Por la ratificación §12-3 del Plan 52 ("si el T0 confirma el defecto gemelo en R2 **con caso real** → plan hijo"), **no se requiere plan hijo**; el riesgo queda declarado como latente (una fuente futura con una sola fila-empresa lo dispararía). El T0 de este plan lo re-verifica por conteos read-only (S1.4); si aparece un caso real → **residuo declarado con dueño** (plan hijo con su T0), nunca fix en este plan.
- **R4 no tiene el defecto**: una sola fila `A=label ∧ B=Total` por `FirstOrDefault`, sin índice de ocurrencia (`ExtraerCeldasR4`, `*.cs:1718-1734`). Nada que declarar salvo que el inventario lo desmienta (entonces STOP).

### D7. E2E 2026101: qué desbloquea este plan y qué queda

**Desbloquea:** el residuo `R-PARCIAL-OCT-ASE3-R1-OCC` (Plan 50 §8) deja de abortar (`TotalMedio=[1]` → throw) y pasa a valores por sección + 0s (oct-1 ASE3 OCCIDENTE: `F237/F216/F242/F222` por contrato). Es el único bloqueo 2026101 que este plan cierra.

**Queda (dueños ajenos, no se tocan):** banco ya cerrado (Plan 51); `R-SIGMA-Q2` (gate Σ acotado Q1); `R-MAPA-V2-R2-ASE4` (A2 del Plan 50, celda fórmula en v2); `ERR-SIN-SLOT` del contrato de sección R1 en fuentes Q1 (Planes 41/46); 20 rojos preexistentes del reemplazo v2 `d55e621`. El e2e 2026101 completo sigue con esos bloqueos tras este plan; prometer su cierre sería falso.

---

## FASE 3 — SPEC

### 6. Contratos conductuales (viejo → nuevo)

| Situación | Hoy | Plan 53 |
|---|---|---|
| Empresa con bloques en secciones S₁<…<Sₖ (Totales y Subs por separado) | `Main=[0]`, `Menor=[count-1]`, `Medio=[1]` por `\|F\|` | Cada operando = fila de su **sección declarada** en el contrato (D1); ausente → `0m`; 2 filas en la misma sección → STOP |
| Única sección de recaudo (oct-1 ASE5) | `Main=[0]` por magnitud (correcto por accidente) salvo D3-del-52 que lo borraba | Contrato por celda: `Main(UnicaRecaudo)` = valor; `Menor`/`Medio` = 0 donde se declaren; **ninguna pérdida** (ENERBIT `F487` intacto) |
| Sección declarada ausente (`Menor`/`Medio`/`SubsMenor`/`Esp*` sin fila) | throw (`ValorDe`/`EspecialesDe`) o duplicado (`count-1==0`) | **0 explícito**, sin throw, sin huella nueva (la huella `FuentesAusentes` solo cubre ausencia total del label — intacta) |
| 1 sola fila oportuna no-primera (oct-1 ASE3) | throw en `TotalMedio` | `Main`=valor, resto 0 (caso §9 #11–#12) |
| 1 sola fila primera-que-es-única (oct-1 ASE5) | valor por accidente | `Main(UnicaRecaudo)`=valor (D5) |
| 2 secciones (19 grupos) | `Medio == Menor` (mismo índice 1) | `Medio(Intermedia)` = 0; `Main`/`Menor` por sección declarada |
| 3 secciones, magnitud≠ordinal (ago-2 ASE2) | `Menor`/`Medio` por magnitud | por **posición de sección** (§9 #6–#7, arbitrado por el manual) |
| 1 sola sección de Subs (`SubsMenor` duplica) | `SubsMenor = [count-1] = SubsMain` | `SubsMenor(Primera)` = 0 (§9 #1, #5, #8; hojas inertes, pero diff values-only) |
| Empate de sección (2 filas-empresa en la misma sección) | colapso silencioso o throw por índice | **STOP** con ASE + empresa + sección + filas (no hay regla de desempate aprobada) |
| Invariantes `Componente` / `Subs/Cont` / `Total` ausentes | fail-fast | **fail-fast intacto** (mensajes y tipo intactos) |
| Ausencia total del label | 0 + huella (Plan 50) | intacto, sin cambios |
| `Block5Total`, columnas 5/11, `CalcularVisibleR1`, writer, plantillas | — | intactos |

### 7. Cambios archivo por archivo

- **T0 (H0, read-only bloqueante — §10):** fixture `SeccionPorOperandoR1T0Tests` (BCL ZIP+XML, cero producción) + doc `Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md`. Congela: (i) matriz del T0-52 re-verificada (52 grupos / 180 filas-empresa, partición/orden `SeccionPorCierre` vs `PorFronteraArriba` fila a fila); (ii) tabla celda → sección esperada (56 entradas, triple evidencia D2 por celda: geometría + `<f>` consumidora + manual); (iii) lista cerrada de Δ (§9: los 17 + los que aparezcan, con `VisibleR1` por caso); (iv) tratamiento `TotalMedio` (D4); (v) inventario R2/R4 (D6). **Sin T0 verde no se abre H1.**
- **H1.1 (Core, nuevo puro): `SeccionEsperadaR1` + `ResolvedorOperandosEmpresaR1`** (`Remuneracion.Core/Models/`, BCL puro, D-H: sin I/O, sin período, sin empresa). `SeccionEsperadaR1` (D1: `UnicaRecaudo/Primera/UltimaOportuna/Intermedia`). `ResolvedorOperandosEmpresaR1`: `IReadOnlyList<(FilaIdx, SeccionPorCierre)> × SeccionEsperada → (fila | ausente | STOP-ambigüedad)`. Implementa EXACTAMENTE D1/D4/D5 sobre secciones ya tipadas (ord primera/intermedia/última-oportuna/única por lista de Totales y de Subs). Sin predicados nuevos: recibe secciones resueltas. + unit tests sintéticos de la regla (ordinales, ausencias, empate→señal STOP).
- **H1.2 (Infrastructure): `ExtraerCeldasR1`** (`ExcelDataReaderWorkbookLeafInputReader.cs:1560-1640`): conserva firma, invariante `indiceComponente`, `EsFilaEmpresa`, `Bloque5Total`, rama Plan-50 de ausencia total y `ValorDe/EspecialesDe` (ahora solo se llaman con fila existente — el throw por índice desaparece por construcción, el método queda como defensa). Cambia SOLO la resolución: guarda `(fila, indice)` en vez de solo `fila`; construye/reutiliza la lista `FilaEspejoR1` del sheet (la misma que alimenta `LeerEspejoR1` — extraer conversor compartido si no existe punto único); tipa cada fila-empresa con `SeccionPorCierre`; delega la elección al resolvedor Core; ausente → `0m`; ambigüedad → throw accionable nombrando ASE+empresa+sección+filas. `OrderByDescending(|F|)` **se elimina** de este camino (es el defecto).
- **H1.3 (Core, contrato): `WorkbookLeafCellMapPorEmpresa.cs:24-128`**: añadir `SeccionEsperadaR1` y extender `EditablesR1PorEmpresa` a `(Celda, FuenteR1, SeccionEsperada)[]` con la tabla del T0 (56 entradas, cada una con comentario de triple evidencia). Reescribir los 9 doc-comments de `FuenteR1` con la semántica por sección (citan `SeccionPorCierre` + este plan). **Cero cambios de miembros, orden o celdas fuera del tercer campo**; `CalcularVisibleR1` (`*.cs:1777-1805`) intacto en lógica (sigue agregando por `FuenteR1`).
- **H1.4 (Tests):** reescritura del throw-test `FuentesAusentesPlan50Tests.S2_1_ParcialConLabelPresente` a aserciones de valor (baja+alta declarada §9; el método permanece, cambia el cuerpo + conserva el comentario de trazabilidad a este plan); tests de no-duplicación con insumos REALES por ruta explícita (`NoDuplicado_Oct1_Ase3_Occidente_UnaFilaOportuna`: `TotalMenor=0 ∧ TotalMain=286.711.907,53 ∧ TotalMenor≠TotalMain`; `NoPerdida_Oct1_Ase5_Enerbit_UnicaRecaudo`: `F487=160.186.886,97 ∧ VisibleR1≥0`); tests de magnitud≠ordinal (ago-2 ASE2 ENEL/OCCIDENTE por posición, con valores del manual); unit tests del resolvedor.
- **H1.5 (gates de cierre):** `git diff --check`, CRLF, SHAs, greps (§8).

### 8. Invariantes y gates de cierre

- Build 0 warnings; CRLF sin mezclar en lo tocado; sin `git commit/push`; SHAs v1/v2 bit-idénticos antes/después (verificar por hash en H2).
- `grep` de cierre: ningún `ElementAtOrDefault`/`[count-1]`/`OrderByDescending` restante en el camino `ExtraerCeldasR1`; ningún `if (empresaId`/`aseId` nuevo (D-H); ningún `<f>` tocado (el reader no escribe fórmulas por construcción); ningún `SeccionEsperada` sin cita del T0.
- Criterios de aceptación: (i) T0 verde con matriz completa + tabla celda→sección con triple evidencia + lista §9 exacta; (ii) H1 implementa D1–D7 sin STOP pendiente; (iii) H2 set-diff == lista §9 exacta (fuera de lista = STOP, sin excepciones); (iv) code-review **APROBADO** (0 bloqueantes); (v) SHAs + `git diff --check` limpios.

---

## FASE 4 — TASKS

### T0 (H0, read-only bloqueante — sin T0 verde no se abre H1)

Fixture: `Remuneracion.IntegrationTests/SeccionPorOperandoR1T0Tests.cs` (solo lectura por rutas explícitas `Docs/Prueba {Julio-1, Julio-2, Agosto-2, Octubre-1}/Insumos/{1-Promoambiental,…}`, nunca `Insumos.Raiz()`). Evidencia: `Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md`.

| ID | Tarea | Hecho en |
|---|---|---|
| S1.1 | Re-verificar la matriz del T0-52 (52 grupos / 180 filas-empresa): conteo `filasTotal`/`filasSubs`, firma A–E completa, valor F, sección por `SeccionPorCierre` Y por `PorFronteraArriba`, frontera que la abre/cierra. Gate de partición/orden: idéntica en ambas lecturas fila a fila; cualquier desacuerdo de partición/orden = STOP | T0 §1 + fixture |
| S1.2 | Tabla celda → sección esperada (56 entradas `EditablesR1PorEmpresa`): por cada celda, triple evidencia D2 (geometría del bloque oficial que la contiene + texto-`<f>` de la fórmula que la consume + valor del manual donde exista). Casos obligatorios: `F237/F216/F222` (OCC ASE3), `F487` (ENERBIT ASE5), `F237→Main` vs `F216→Menor`, `F483/F474` (ENEL ASE5), `F29/F242` (SubsMenor inertes), ASE1/ASE4 completos | T0 §2 + fixture |
| S1.3 | Casos cardinales: oct-1 ASE5 (3 empresas, 1 fila en primera-única → `Main(UnicaRecaudo)`); oct-1 ASE3 OCCIDENTE (solo oportuna → `Main` + 0s); ago-2 ASE2 ENEL/OCCIDENTE (magnitud≠ordinal → posición); jul-2 ASE5 ENEL (1 fila no-primera → `Main` + `Menor=0`); 2 filas (19 grupos → `Medio=0`); 3 filas (18 grupos) | T0 §3 |
| S1.4 | Inventario R2/R4 read-only (D6): conteos por label en los 4 períodos. Si aparece 1-fila-empresa en R2 o multi-fila en R4 → **residuo declarado con dueño** (plan hijo con su T0), nunca fix aquí | T0 §4 |
| S1.5 | Congelar la lista cerrada de Δ: los 17 del T0-52 §6 + los que aparezcan al aplicar el contrato por celda (con `VisibleR1` por caso y arbitraje del manual §8). Todo Δ fuera de esta lista en H2 = STOP | T0 §5; fija §9 |
| S1.6 | Gate de tipabilidad: toda celda del contrato tipable con la triple evidencia D2; **si alguna no → STOP y reporte** (no inventar). `TotalMedio` con veredicto D4 (2 Δ declarados, hojas inertes) | T0 §6 |
| S1.7 | Reconciliación de etiquetas (D3): citar Plan 46 §2.2; demostrar partición/orden idénticas en 180/180; reescribir la justificación D1-del-52 sin cambiar el veredicto funcional | T0 §7 |

### H1 (implementación contract-first, traza a §6–§7)

| ID | Tarea | Traza |
|---|---|---|
| H1.1 | `SeccionEsperadaR1` + `ResolvedorOperandosEmpresaR1` puros + unit tests (roles, ausencias→ausente, empate→señal STOP, única→Main) | D1/D4/D5, §7 |
| H1.2 | Swap en `ExtraerCeldasR1`: `(fila,indice)` + `SeccionPorCierre` reutilizando construcción `FilaEspejoR1` + resolvedor; ausente→0; ambigüedad→throw accionable; retirar `OrderByDescending(\|F\|)` del camino | A1/A2, §7 |
| H1.3 | Contrato `EditablesR1PorEmpresa` a 3-tupla con la tabla del T0 (56 entradas + comentarios de triple evidencia) + doc-comments `FuenteR1` reescritos | D1/D2, A1 |
| H1.4 | Reescribir `S2_1_ParcialConLabelPresente` a aserciones de valor (baja+alta declarada §9) + tests de no-duplicación / no-pérdida / magnitud≠ordinal con insumos REALES (desigualdades explícitas `Menor≠Main`) | A2/D5, §9 |
| H1.5 | `git diff --check`, CRLF, SHAs, greps de cierre (§8) | §8 |

### H2 (regresión y evidencia)

Suite completa → `TestResults/p53-h2-suite.trx` + `TestResults/p53-h2-setdiff.txt` contra `p51-h2-suite.trx` (1260/1193/67/0). PASS criterios §8-iii. Code-review independiente (subagente `code-reviewer`, este plan como spec): **APROBADO** requerido (0 bloqueantes).

### H3 (docs)

Párrafo de doctrina en `.opencode/project-context.md` (contrato por operando R1: tabla de roles `UnicaRecaudo/Primera/UltimaOportuna/Intermedia` resumida, predicado `SeccionPorCierre`, remisión al T0); nota de supersesión en `plans/52` (§13 Ruta B → este plan); este plan queda como fuente de verdad. Sin cambios al Instructivo (la huella `FuentesAusentes` se reutiliza, no hay huella nueva).

### 9. Lista cerrada de altas/bajas/Δ esperados (H2 la verifica exacta)

Adoptada del T0-52 §6 (17 operandos-celda en 13 grupos; 7 mueven el `VisibleR1`). El T0 de este plan (S1.5) la confirma o la amplía **antes** de H1; lo que sigue es la base mínima — todo Δ fuera de la lista final del T0 = STOP:

| # | Período | ASE | Empresa | Celda | Operando | Hoy | Plan 53 | Δ |
|---|---|---|---|---|---|---|---|---|
| 1 | Jul-2 | 1 | OCCIDENTE | F29 | SubsMenor | 168.024.399,69 | 0 | −168.024.399,69 |
| 2 | Jul-2 | 3 | OCCIDENTE | F242 | SubsMenor | 404.748.758,01 | 0 | −404.748.758,01 |
| 3 | Jul-2 | 3 | OCCIDENTE | F222 | TotalMedio | −16.458,46 | 0 | +16.458,46 |
| 4 | Jul-2 | 5 | ENEL | F474 | TotalMenor | 12.823.196.468,26 | 0 | −12.823.196.468,26 |
| 5 | Ago-2 | 1 | OCCIDENTE | F29 | SubsMenor | 206.290.801,38 | 0 | −206.290.801,38 |
| 6 | Ago-2 | 2 | ENEL | F83 | TotalMenor | 131.543,19 (mag) | −1.414.859,14 (posición f7) | −1.546.402,33 |
| 7 | Ago-2 | 2 | OCCIDENTE | F89 | TotalMenor | 24.890,00 (mag) | −28.650,84 (posición f12) | −53.540,84 |
| 8 | Ago-2 | 3 | OCCIDENTE | F242 | SubsMenor | 360.488.591,23 | 0 | −360.488.591,23 |
| 9 | Ago-2 | 3 | OCCIDENTE | F222 | TotalMedio | −522,80 | 0 | +522,80 |
| 10 | Ago-2 | 5 | ENEL | F474 | TotalMenor | 13.260.342.962,39 | 0 | −13.260.342.962,39 |
| 11 | Oct-1 | 3 | OCCIDENTE | F216 | TotalMenor | 286.711.907,53 | 0 | −286.711.907,53 |
| 12 | Oct-1 | 3 | OCCIDENTE | F242 | SubsMenor | 145.489.038,47 | 0 | −145.489.038,47 |
| 13 | Oct-1 | 5 | ENEL | F483 | TotalMain | 7.414.210.944,89 | 7.414.210.944,89 (`UnicaRecaudo`) | 0 |
| 14 | Oct-1 | 5 | ENEL | L483 | EspecialesMain | 10.471.696,87 | 10.471.696,87 (`UnicaRecaudo`) | 0 |
| 15 | Oct-1 | 5 | ENERBIT | F487 | TotalMain | 160.186.886,97 | 160.186.886,97 (`UnicaRecaudo`) | 0 |
| 16 | Oct-1 | 5 | OCCIDENTE | F497 | TotalMain | 134.612.117,14 | 134.612.117,14 (`UnicaRecaudo`) | 0 |
| 17 | Oct-1 | 5 | OCCIDENTE | L497 | EspecialesMain | 1.908.241,19 | 1.908.241,19 (`UnicaRecaudo`) | 0 |

Lectura de la tabla (cambio deliberado vs Plan 52 §9): las filas #13–#17 **no son Δ** — son la corrección del error D3-del-52. El contrato por celda (D1/D5) las fija en `UnicaRecaudo → Main`; el T0 (S1.3) lo demuestra con la geometría de ASE5-oct (2 únicas secciones: `Oportuno#1` de recaudo + bloque de Subs `Oportuno#2`). Si el T0 encontrara una celda fuera de estas 17 con Δ → se añade a la lista **antes** de H1 con evidencia; después de H1, cualquier Δ fuera de lista = STOP.

- **Altas**: fixture T0 (~N tests S1.x) · unit tests resolvedor (~M) · 3 no-duplicación/no-pérdida/magnitud (REAL, §7-H1.4) · 1 reescritura S2_1 (baja+alta declarada).
- **Bajas**: `S2_1_ParcialConLabelPresente_SigueSiendoFailFast` en su forma throw (reemplazada, no eliminada: el método permanece con aserciones de valor).
- **Cambios de valor esperados**: los Δ #1–#12 (con arbitraje del manual §8 para #4, #6, #7, #10); #13–#17 identidad blindada (ante diff se revierte el pase, no el golden).

---

## 10. Índice de evidencia (a producir)

| Evidencia | Dónde |
|---|---|
| Matriz re-verificada (52 grupos / 180 filas-empresa) + tabla celda→sección (56 entradas, triple evidencia) | `Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md` §§1–2 + fixture `SeccionPorOperandoR1T0Tests` |
| Fuentes R1 (20 archivos) | `Docs/Prueba Julio-1/Insumos/…` · `Prueba Julio-2/…` · `Prueba Agosto-2/…` · `Prueba Octubre-1/…` (`Recaudoporcomponente_*.xlsx`, `Sheet1`) |
| Oráculo manual | `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx!REMUNERACION_ENEL!D13` y análogas (T0-52 §8); `CONSOLIDADO_TOTAL RECAUDO!D47:D51` (geometría legacy, R-TRAMPA-ANCLA) |
| Plantilla oficial v2 (slots + `<f>` consumidoras) | `Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx!Reporte Componentes R1` (solo lectura) |
| Inventario R2/R4 | T0 §4 (conteos sobre `RerpoteDetalleSaldosaFavor_*` y `ReversiónPorComponente_*` de los 20 ASE×período) |
| Regresión | `TestResults/p53-h2-suite.trx` + `TestResults/p53-h2-setdiff.txt` vs `p51-h2-suite.trx` (1260/1193/67/0) |

---

## 11. Riesgos

| ID | Riesgo | Mitigación |
|---|---|---|
| R1 | Una celda del contrato no es tipable con la triple evidencia (geometría ambigua, `<f>` sin rol claro, sin manual) | T0 S1.6: **STOP y reporte** (encargo). No se inventa ningún rol |
| R2 | 2 filas-empresa en la MISMA sección (empate sin regla aprobada) | STOP por diseño (H1.1 lo señala con ASE+empresa+sección+filas, no lo resuelve) |
| R3 | Filas-empresa con `E` distinto de vacío/concepto no-`Total` dentro de la misma sección | S1.1 registra `E`; si aparece → STOP (el operando exige concepto `Total`) |
| R4 | Una forma futura trae 4+ secciones o una sección `Aplicacion` donde el contrato espera `Oportuno` | T0 nuevo (re-congelamiento citado); nunca rama ad hoc. El resolvedor solo conoce roles relativos |
| R5 | Tentar "arreglar" columnas 5/11 o el header-scan de componentes al tocar el reader | Prohibido por §4.4; el grep de cierre lo vigila |
| R6 | R2 confirma un caso 1-fila real en el inventario S1.4 (duplicado `Total==SubsCont` latente) | Se declara como residuo con dueño (plan hijo con su T0, ratificación 52 §12-3); no expande este plan |
| R7 | El revisor confunde `UnicaRecaudo → Main` con un fallback posicional | El contrato es por celda declarada con triple evidencia, no por conteo (`count==1 ⇒ Main` prohibido como regla global; solo vale donde el T0 lo congela celda por celda) |

## 12. Decisiones abiertas para el Ingeniero (el plan NO está aprobado)

1. **Modelo del contrato (D1):** ¿se ratifica (a) tupla `(Celda, Fuente, SeccionEsperada)` con enum `UnicaRecaudo/Primera/UltimaOportuna/Intermedia` frente a (b)/(c)? Sin esto no se abre H0 (el T0 congela la tabla en esa forma).
2. **Derivación de la sección esperada (D2):** ¿se ratifica la triple evidencia (geometría + `<f>` consumidora + manual) y que la celda no tipable dispara STOP?
3. **Predicado (D3):** ¿se ratifica `SeccionPorCierre` como tipador único y la reescritura de la premisa D1-del-52 (coincidencia de partición/orden, no de etiqueta, citando Plan 46 §2.2)?
4. **`TotalMedio` (D4):** ¿se ratifica `Intermedia-ssi-≥3-si-no-0` con los 2 Δ como divergencia declarada con dueño (extensión de la ratificación 52 §12-1 a `SubsMenor`/inertes)?
5. **Lista cerrada (D5/§9):** ¿se ratifica la base de 17 filas (#1–#12 con Δ, #13–#17 identidad blindada `UnicaRecaudo`) y que el T0 la amplíe antes de H1 pero H2 la exija exacta?
6. **R2/R4 (D6):** si S1.4 confirma el defecto gemelo en R2 con caso real, ¿se autoriza plan hijo con su T0 (recomendado) o se declara deuda sin plan?
7. **Nada más:** D7 y el método T0→H1→H2→H3 vienen cerrados del encargo; no requieren debate salvo que el T0 dispare un STOP.

**Sin estas aprobaciones el plan no avanza de BORRADOR a APROBADO, y sin APROBADO no se abre H0.**

**RATIFICACION 2026-10-22 (Ingeniero) — APROBADO 7/7:** (1) modelo (a) `(Celda, Fuente, SeccionEsperada)` con enum `UnicaRecaudo/Primera/UltimaOportuna/Intermedia`; (2) triple evidencia geometría + `<f>` + manual, celda no tipable → STOP; (3) `SeccionPorCierre` como tipador único (premisa D1-del-52 reescrita); (4) `TotalMedio = Intermedia si ≥3, si no 0` con los 2 Δ como divergencia declarada; (5) lista §9 de 17 filas cerrada por el T0 antes de H1; (6) R2 hijo solo con caso real; (7) nada más. **H0 HABILITADO.**

---

## Architecture Validation Certificate

- **SOLID:** SRP (T0 = evidencia/matriz + tabla celda→sección como DATOS; `SeccionEsperadaR1`/`ResolvedorOperandosEmpresaR1` = juicio puro en Core; reader = evidencia + 0 explícito; `ValidadorBasico` intacto; writer no juzga); OCP (nueva forma/sección = fila citada en T0 nuevo, cero ramas; fuente ausente nueva = cero código por Plan 50); LSP/DIP (cambios vía interfaces existentes `IWorkbookLeafInputReader`/`IValidador`/`IPlantillaWriter`; campos aditivos, contratos intactos); ISP (contratos `IRecaudoReader`/`IPlantillaWriter` intactos; API `PorFronteraArriba` intacta para el consumidor Plan 35-T2).
- **Best Practices:** contract-first (H0 congela celda→sección + lista Δ antes del juicio); fail-fast accionable (ASE + empresa + sección + filas; invariantes T0e intactos en formato); gates que leen el workbook generado (texto-`<f>` + literales, nunca `<v>` stale) + dominio por firma ±0.5; values-only; D-H (cero literales/ramas por período o empresa en producción — los roles son relativos, no literales); CRLF sin mezclar; sin commits.
- **Performance:** tipado por cierre una vez por fila-empresa (O(filas del bloque) por ASE); agrupación por (firma, sección) con cursor por grupo; resolución por operando O(1) sobre grupos en memoria; lectura acotada en H0 (BCL ZIP+XML); regens a temp fresca sin copiar de más.

---

## 13. Cierre (2026-10-22)

- **Ejecución:** H0 VERDE (T0 `Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md` + fixture 11 verdes) · H1 DONE (contrato 56 entradas + resolvedor puro + swap en `ExtraerCeldasR1`) · H2 VERDE (`p53-h2-suite.trx` = **1289/1222/67/0**; set-diff vs `p51-h2-suite.trx`: **29 altas del plan, 0 bajas, 0 rojos nuevos**) · H3-docs ejecutado.
- **Code-review independiente:** **APROBADO CON OBSERVACIONES** (0 bloqueantes). Verificado: contrato fiel al T0 (56 entradas); el ordinal eliminado del camino; oct-1 ASE5 ENERBIT intacto; los 2 Δ son la eliminación exacta de un doble conteo (aritmética al centavo); `UnicaRecaudo ≡ UltimaOportuna` correcto (rol derivado, no fallback); tests no debilitados; archivos extra son destructuring mínimo.
- **Desviaciones ratificables (requieren ratificación del Ingeniero):**
  1. **D2.1** — 34/56 celdas sin geometría oficial propia: condición **preexistente** (Plan 41 E1-2/E1-4; T0-52 §8-b/c). Tipabilidad sostenida por fuente (180/180) + `<f>` consumidora + manual (4/4 al centavo) + identidad julio-1. Deuda declarada con dueño (frente E1-2).
  2. **Δ #11–#12 (lista §9 → 19 filas)** — `F474` (oct-1 ASE5 ENEL) y `F477` (OCCIDENTE) a 0; `VisibleR1` ENEL `14.647.388.373,02→7.233.177.428,13`, OCC `306.177.388,95→171.565.271,81`. Es la eliminación del doble conteo `Menor=[count-1]==Main` en bloque de sección única; ampliación hecha por el T0 **antes** de H1.
- **Observaciones menores (declaradas, no bloquean):** W-5 (`F86` Block5Total declara `UltimaOportuna` sin consumidor en producción), W-6 (contador del script de set-diff con −1 simétrico), W-7 (corregido: blank line at EOF), W-8 (errata de SHA en T0 §10), S-1 (corregido: "53 entradas" → 56).
- **Deuda preexistente registrada (no del 53):** `GoldenInterventoriaTests.CapaA_InterventoriaYLEspeciales_AmbosCanonicos` rojo por `ERR-SIN-SLOT` (frente Slot R1, Planes 41/49).
- **RATIFICACION 2026-10-22 (Ingeniero):** ratificadas las 2 desviaciones (eliminación del doble conteo en oct-1 ASE5 y D2.1 preexistente). **Plan 53 CERRADO sin condiciones.** El e2e 2026101 desbloquea `R-PARCIAL-OCT-ASE3-R1-OCC`; el resto sigue con sus dueños.
