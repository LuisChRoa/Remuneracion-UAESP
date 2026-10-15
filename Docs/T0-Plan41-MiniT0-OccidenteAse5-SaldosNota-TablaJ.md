# T0 (Plan 41) — Mini-T0: evidencia congelada desde bytes (Occidente ASE5 · saldos-nota · tabla J)

> **Frente:** Plan 41 (§0/T0, R-T-1). **Naturaleza:** T0 DIAGNÓSTICO — **cero código de producción**.
> Lectura SOLO con BCL (`System.IO.Compression.ZipArchive` + `System.Xml.Linq`), sin Excel/COM, sin
> OpenXml SDK. **El OFICIAL nunca se modifica** (SHA256 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214`).
> **Fixture read-only:** `Remuneracion.IntegrationTests/T0FrenteEspejoR1Ase5EnelTests.cs` (extendido; congela
> los bytes citados aquí; no toca producción).

## 0. Insumos (rutas explícitas)

| Insumo | Ruta |
|---|---|
| Oficial | `Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx` |
| Manual oráculo | `Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx` |
| Salida app (defecto) | `Docs/Prueba Julio-2/Resultado/Resultado5/Remuneración 202607-2 Total.xlsx` |
| Fuente R1 ASE5 | `Docs/Prueba Julio-2/Insumos/5-Área Limpia/Recaudoporcomponente_…_20268465657830.xlsx` |
| Fuente R1 ASE2 Q1 | `Docs/Prueba Julio-1/Insumos/2-Lime/Recaudoporcomponente_…_202671616523961.xlsx` |

---

## 1. OCCIDENTE ASE5 — divergencia de ESTRUCTURA de fórmula (F700 vs F578)

### 1.1 Textos-`<f>` exactos (byte-fiel)

| Documento | Celda | Texto-`<f>` |
|---|---|---|
| Oficial | `Reporte Componentes R1!F700` | `F647+F673+F612-L612-L647` (3 términos F, 2 L) |
| Manual | `Reporte Componentes R1!F578` | `F530+F551-L530` (2 términos F, 1 L) |

### 1.2 Operandos y valores (oficial en su geometría; manual en la suya)

Oficial — bloque OCCIDENTE ASE5 (secciones: 612=1.ª, 628=Oportuno#1, 647=Aplic#1, 659=Oportuno#2, 673=Aplic#2):

| Celda oficial | Firma (C|D) | F | L | Sección (frontera arriba) |
|---|---|---|---|---|
| `F612`/`L612` | `OCCIDENTE|Total` | 0 | 0 | 1.ª (sin frontera) |
| `F628`/`L628` | `OCCIDENTE|Total` | 0 | 0 | Oportuno #1 (Mes 613) |
| `F647`/`L647` | `OCCIDENTE|Total` | 0 | 0 | Aplicación #1 (Aplicacion 629) |
| `F659`/`L659` | `OCCIDENTE|Total` | 0 | 0 | Oportuno #2 (Mes 648) |
| `F673`/`L673` | `OCCIDENTE|Total` | 0 | 0 | Aplicación #2 (Aplicacion 660) |

Manual — bloque OCCIDENTE ASE5 (legacy): `F511`=417.519,62 (1.ª), `F530`=229.202.451,62 / `L530`=5.904.652,23 (Aplic#1),
`F537`=24.329,38 (Oportuno), `F551`=81.691.821,42 (Aplic#2).

Fuente R1 ASE5 (julio-2) — filas OCCIDENTE y su sección:

| Fila fuente | C | F | L | Sección (frontera arriba) |
|---|---|---|---|---|
| 13 | OCCIDENTE | 417.519,62 | 0 | 1.ª (antes de Aplicacion 14) |
| 32 | OCCIDENTE | 229.202.451,62 | 5.904.652,23 | Aplicación #1 (Aplicacion 14) |
| 39 | OCCIDENTE | 24.329,38 | 0 | Oportuno #1 (Mes 33) |
| 53 | OCCIDENTE | 81.691.821,42 | 0 | Aplicación #2 (Aplicacion 40) |

### 1.3 Evaluación

Colocación por sección: src13→612, src32→647, src39→628, src53→673.

```
Oficial F700 = F647 + F673 + F612 - L612 - L647
            = 229.202.451,62 + 81.691.821,42 + 417.519,62 - 0 - 5.904.652,23
            = 305.407.140,43
Manual  F578 = F530 + F551 - L530
            = 229.202.451,62 + 81.691.821,42 - 5.904.652,23
            = 304.989.620,81
Δ = 417.519,62 = F612 (la fila de la 1.ª sección de OCCIDENTE)
```

### 1.4 VEREDICTO DE NEGOCIO (motivado)

- El **oficial es INTERNAMENTE CONSISTENTE**: sus 5 anclas OCCIDENTE tienen la MISMA estructura de 3
  términos, incluyendo el término de la 1.ª sección:
  `F95 = F69+F46+F15-L15-L46`, `F274 = F207+F170+F119-L119-L170`, `F435 = F352+F329+F298-L298-L329`,
  `F588 = F547+F510+F459-L459-L510`, `F700 = F647+F673+F612-L612-L647`.
- El **manual es INTERNAMENTE INCONSISTENTE**: sus anclas OCCIDENTE ASE1–4 tienen la MISMA estructura de
  3 términos (`F73 = F47+F31+F11-L11-L31`, `F226 = F159+F131+F93-L93-L131`,
  `F363 = F280+F264+F243-L243-L264`, `F488 = F447+F421+F383-L383-L421`), pero su ASE5 tiene 2 términos
  (`F578 = F530+F551-L530`), OMITIENDO el término de la 1.ª sección que sus propias ASE1–4 incluyen.
- Conclusión: la **anomalía es la ancla ASE5 del MANUAL**, no el oficial. Pero el valor del manual
  (304.989.620,81) NO es alcanzable por el oficial en esa geometría, y el valor del oficial
  (305.407.140,43) NO es el del manual. Es una **divergencia estructural de plantilla/manual** que la
  UAESP debe conocer y arbitrar.
- **NO se resuelve por supuesto** (D-E): no se edita el oficial (opción C prohibida), no se "ajusta" el
  gate, y el par **queda DECLARADO** (nunca verde) hasta que el negocio fije la estructura correcta.

**Hallazgo para la UAESP:** la fórmula `F700` del oficial incluye la fila de la 1.ª sección de OCCIDENTE
(`+F612-L612`), que la fórmula `F578` del manual omite; con la fuente julio-2 la diferencia material es
de 417.519,62. La UAESP debe confirmar si el "Recaudo Oportuno" de OCCIDENTE incluye o no la 1.ª sección.

---

## 2. SALDOS POR NOTA / RETRIBUCION NEGATIVA — geometría del bloque oficial

### 2.1 Geometría oficial (6 filas uniformes por ASE)

`SALDOS POR NOTA` (col B = concepto; col C = Total; col I = Especiales):

| ASE | Filas | Conceptos (fila → concepto) | Visible Cn-In |
|---|---|---|---|
| 1 | 3–8 | 3 Vlr Servicio · **4 Vlr Intereses** · 5 Componente · 6 Subsidio · 7 Subs/Cont · **8 Total** | `C10 = C8-I8` |
| 2 | 16–21 | 16 Vlr Servicio · **17 Vlr Intereses** · 18 Componente · 19 Subsidio · 20 Subs/Cont · **21 Total** | `C23 = C21-I21` |
| 3 | 29–34 | 29 Vlr Servicio · **30 Vlr Intereses** · 31 Componente · 32 Subsidio · 33 Subs/Cont · **34 Total** | `C36 = C34-I34` |
| 4 | 42–47 | 42 Vlr Servicio · **43 Vlr Intereses** · 44 Componente · 45 Subsidio · 46 Subs/Cont · **47 Total** | `C49 = C47-I47` |
| 5 | 55–60 | 55 Vlr Servicio · **56 Vlr Intereses** · 57 Componente · 58 Subsidio · 59 Subs/Cont · **60 Total** | `C62 = C60-I60` |

`RETRIBUCION NEGATIVA` (mismas filas por ASE; el orden de conceptos cambia: `Vlr Intereses` PRIMERO):

| ASE | Filas | Conceptos | Visible Cn-In |
|---|---|---|---|
| 1 | 3–8 | **3 Vlr Intereses** · 4 Vlr Servicio · 5 Componente · 6 Subsidio · 7 Subs/Cont · **8 Total** | `C10 = C8-I8` |
| 2 | 16–21 | **16 Vlr Intereses** · 17 Vlr Servicio · 18 Componente · 19 Subsidio · 20 Subs/Cont · **21 Total** | `C23 = C21-I21` |
| 3 | 29–34 | **29 Vlr Intereses** · 30 Vlr Servicio · 31 Componente · 32 Subsidio · 33 Subs/Cont · **34 Total** | `C36 = C34-I34` |
| 4 | 42–47 | **42 Vlr Intereses** · 43 Vlr Servicio · 44 Componente · 45 Subsidio · 46 Subs/Cont · **47 Total** | `C49 = C47-I47` |
| 5 | 55–60 | **55 Vlr Intereses** · 56 Vlr Servicio · 57 Componente · 58 Subsidio · 59 Subs/Cont · **60 Total** | `C62 = C60-I60` |

> El bloque OFICIAL trae SIEMPRE la fila `Vlr Intereses` (6 filas). El mapa de la app
> (`WorkbookLeafCellMapAjustesSfT`) estaba congelado de la plantilla canónica LEGACY (5/6 filas
> dispares) → desplazaba las filas fuente +1 (escribía el `Total` de la fuente en `Subs/Cont`).
> Plan 41/T2 re-congela el mapa a esta geometría oficial.

### 2.2 Geometría legacy del manual (oráculo del administrativo)

`SALDOS POR NOTA` (legacy): ASE1 filas 3–7 (SIN `Vlr Intereses`), ASE2 15–20, ASE3 28–32, ASE4 40–45,
ASE5 53–58. El `Total` de ASE1 (manual) está en `C7 = 973.693,46` y el visible es `C9 = C7-I7`.

**Diferencia estructural:** el manual (legacy) NO trae la fila `Vlr Intereses` en ASE1/ASE3; el oficial
sí. La correspondencia es por (concepto, columna), NUNCA por dirección entre plantillas (R-TRAMPA-ANCLA).

### 2.3 Mapeo fuente→bloque (Plan 41/T2)

El `Total` de la fuente (`SaldosaFavorAplicadosPorNotas_*`) debe caer en la fila `Total` OFICIAL:
`C8` (ASE1), `C21` (ASE2), `C34` (ASE3), `C47` (ASE4); ASE5 = 0 (fuente vacía). El visible resultante
`C10/C23/C36/C49 = Total − Especiales` alimenta `AJUSTES - SF-T!D9/D10/D11/D12` y `D47 = D9+D28`.

---

## 3. INTERVENTORIA — tabla J / `R26:R31` (auto-chequeo `S26:S31`)

### 3.1 Textos-`<f>` (oficial, byte-fiel)

```
Q26 = ='CONSOLIDADO_TOTAL RECAUDO'!L104      (Q27=L105 … Q31=L109)
R26 = =J15   <f t="shared" ref="R26:R31" si="0"/>   (followers R27..R31 sin texto → J16..J20)
S26 = =Q26=R26   (S27 = Q27=R27; S28:S31 shared ref=S28:S31 si=2)
```

### 3.2 Tabla J y variantes (oficial)

| Celda | Valor | Naturaleza |
|---|---|---|
| `J15:J19` | 189.185.988, 266.580.256, 154.788.536, 120.391.083, 128.990.446 | literales (tabla anual) |
| `J20` | 859.936.309 | literal (total) |
| `M26:M30` (2.ª quincena) | = `J15:J19` | declarado Q2 |
| `N26:N30` (1.ª quincena) | 189.185.987, 266.580.255, 154.788.535, 120.391.083, 128.990.445 (`= M−1`) | declarado Q1 |

`CONSOLIDADO!L9:L13 = INTERVENTORIA!M26:M30` (oficial, Q2) → Phase 0 selecciona `N26:N30` en Q1
(R-INTERV-Q). Y `L104:L109` arranca en `L9:L13` → `Q26:Q31` sigue la selección por quincena.

### 3.3 VEREDICTO (motivado)

- **Q2:** `Q26 = L104 = M26 = 189.185.988` y `R26 = J15 = 189.185.988` ⇒ `S26 = TRUE`.
- **Q1:** `Q26 = L104 = N26 = 189.185.987` y `R26 = J15 = 189.185.988` ⇒ `S26 = FALSE` (divergencia real).
- El oficial **NO trae una variante N de la tabla J** (los valores Q1 viven solo en `N26:N30`).
  ⇒ **NO es invariante**: Phase 0 debe reescribir el maestro `R26` por `Periodo.NumeroQuincena`
  (Q2 → `J15` identidad; Q1 → `N26`, que con el shared relativo produce `R26:R31 = N26:N31`).
- **Prohibido sellar TRUE** donde Q1 es FALSE real. La regla la decide el DOMINIO, nunca el contenido.

**Hallazgo para la UAESP:** el auto-chequeo `S26:S31` del oficial está calibrado para Q2 (tabla J = los
valores de la 2.ª quincena); en Q1 evaluaría FALSE por construcción hasta que Phase 0 aplique la variante
de la quincena.

---

## 4. Método y reproducibilidad

- Lectura BCL ZIP+XML del oficial, el manual, la salida de la app y las fuentes R1 por ruta explícita.
- Los textos-`<f>` son el oráculo; los cachés `<v>` solo documentan la variante (nunca la demuestran).
- **Prohibido** `Docs/Insumos` raíz; **prohibido** usar cachés `<v>` como oráculo.
- El oficial se verifica intacto por SHA256 (§0) antes/después.
