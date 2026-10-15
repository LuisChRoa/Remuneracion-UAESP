# Plan 43 — Cierre de observaciones de la revisión independiente (Planes 40/41/42)

> **Alcance:** cerrar los hallazgos de la revisión independiente de código sobre los Planes 40/41/42 (veredicto: **APROBADO CON OBSERVACIONES**, sin hallazgos BLOQUEANTES ni MAYORES en producción) y una observación propia del orquestador, para dejar el aplicativo listo para las **pruebas manuales del Ingeniero**. Cambio esperado: **mínimo y acotado** (una vía de arranque que falla cerrada, renormalización de fin de línea, huella de quién/canal en el reporte, y marca de "sin validar" en el archivo de una corrida fallida). Cero cambios de semántica de negocio, de lectura R1/R2/R4, de cálculo CONSOLIDADO, de la doctrina de divergencia declarada, del criterio estricto, ni del espejo R1.
> **Estado:** APROBADO por el Ingeniero 2026-10-14 ("Sí, proceder con los 4"). EN EJECUCIÓN.
> **Fecha:** 2026-10-14
> **Nota de forma:** plan de cierre compacto, ratificado en conversación; su contenido es el registrado aquí.

---

## 0. Contexto

La revisión independiente (read-only) de los Planes 40/41/42 confirmó: la puerta de **divergencia declarada NO es un coladero** (solo clasifica lo evaluable desde lo escrito; cadena no-resoluble o snapshot sin evidencia nunca se declara; `HayNoDeclaradas` suma los sub-bloques no-declarados y el resto de gates; el filtro `OfreceReconocimiento` y el constructor defensivo del diálogo impiden abrir la superficie en estado mixto); **ningún test fue debilitado** (0 `Skip=`, 0 `Assert.True(true)`, 0 `NotImplementedException`; set-diff de los 39 fallidos idéntico entre trx); los 8 rojos de Q1 siguen rojos por `ERR-SIN-SLOT` sin silenciar; los invariantes duros se cumplen con prueba; build 0 warnings; oficial SHA256 intacto.

Quedaron **tres hallazgos MENORES** y **una observación propia** que este plan cierra.

---

## 1. Unidades y criterios de aceptación

| ID | Qué | Criterio de aceptación |
|---|---|---|
| **O-1** | **Cerrar la trampa de arranque (MENOR 1 de la revisión).** El constructor por defecto de `Form1` compone `ProcesadorPeriodo` **sin oráculo** (8.º argumento `null`) y con emisor de reporte como 9.º: por esa vía la clasificación nunca corre y un run **cerraría en verde sin validar** (W-1 no se ejecuta). Producción usa `Program.cs` (correcto), pero la composición muerta es una trampa si cambia el entrypoint. | Esa vía **no puede** producir un cierre verde sin validación: **falla en duro** (throw documentado) o inyecta el oráculo. Test que lo pruebe en ambos sentidos. |
| **O-2** | **Fin de línea (NIT 4).** 7 archivos nuevos quedaron en LF puro (`SeccionEspejoR1.cs`, `ResolutorValorValidacion.cs`, `SelladoCacheFormulaT2Tests.cs`, `TablaJS26S31T3Tests.cs`, `ValidacionTotalEvaluaT1Tests.cs`, `T0-Plan40-W1-Evalua-Sella.md`, `T0-Plan41-MiniT0-OccidenteAse5-SaldosNota-TablaJ.md`). | 0 archivos con LF puro en el diff; CRLF sin mezclar en todo lo tocado. |
| **O-3** | **Huella de quién y por qué canal (MENOR 2).** El reporte dice "el usuario reconoció explícitamente" sin identidad, y el log no registra canal. El entorno sí lo permite. | El reporte durable y el log registran **quién** aceptó (usuario del sistema) y **por qué canal** (UI o CLI), además de lo ya existente. |
| **O-4** | **Archivo de corrida fallida (observación del orquestador).** Cuando la corrida falla, el writer ya escribió el `.xlsx` antes del oráculo: queda un archivo **con el nombre definitivo** en la carpeta de salida, sin reporte. Riesgo de que alguien lo tome por bueno. Decisión ratificada: **marcarlo como "sin validar"**. | Una corrida que **no** cierra bien **nunca** produce el nombre definitivo: el archivo queda con una marca visible y determinista de "sin validar" (sufijo explícito o subcarpeta de descartes, documentado). Solo una corrida cerrada bien produce el nombre limpio. Prueba en **ambos sentidos**. Los tests que hoy leen la salida de una corrida fallida se **retargetean** a la ruta marcada **con justificación citada**, sin debilitar asserts. |

---

## 2. Invariantes (heredados; TODOS vigentes)

Los de los Planes 40/41/42, sin excepción: los cachés `<v>` de fórmulas **no** son oráculo (el gate evalúa, el writer sella por read-back); criterio estricto para lo **no** declarado (E1-1..E1-5 del Plan 40); divergencia declarada = datos citados con dueño y evidencia, última puerta, doble huella obligatoria y fail-closed; **prohibido el fallback por cursor** (E1-1 del Plan 41) y **prohibido sellar/apagar/re-etiquetar** para que algo pase; solo `<v>`, jamás `<f>` fuera de Phase 0; NUNCA `InsertRow`/`DeleteRow`; tolerancia ±0.5; cero literales/ramas por período o empresa en el diff de producción (D-H); el archivo OFICIAL nunca se modifica (SHA256 `3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214` antes/después); build 0 warnings; CRLF sin mezclar; NUNCA `git commit`/`git push`; rutas explícitas y regens a temp FRESCO; quincena por dominio, nunca por contenido; **los tests NO se debilitan**.

## 3. Gates de cierre (evidencia fresca)

1. `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
2. Focused: los tests de los Planes 40/41/42 + los nuevos de O-1/O-4 → verdes.
3. Suite completa con trx: **línea base 39 Failed / 873 Passed / 912 Total** — set-diff con **cero rojos nuevos**; los **8 de Q1** (declarados por E2-5 del Plan 41) siguen rojos por su causa, ni tocados ni silenciados.
4. Greps: 0 archivos con LF puro en el diff; 0 recursos de UI inventados; 0 literales/ramas por período o empresa; 0 `InsertRow`/`DeleteRow`; 0 `<f>` fuera de Phase 0.
5. Oficial SHA256 intacto antes/después.

## 4. Fuera de alcance (follow-ups citados, no de este plan)

- **Hardening del matcher** de divergencias declaradas (fijar montos absolutos o un segundo discriminador): un defecto interno con Δ exactamente 417.519,62 podría clasificarse declarado. Se aborda al añadir la próxima entrada del catálogo.
- **Contrato de sección Q1** (plan hijo; los 8 rojos declarados).
- **Veredicto de negocio de OCCIDENTE ASE5** (UAESP; la entrada del catálogo se retira por su condición de retiro).
- `.atl/` (caché de tooling) **no debe commitearse**.
