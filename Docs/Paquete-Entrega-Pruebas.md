# Paquete de Entrega a Pruebas — Remuneración Quincenal UAESP

> **Propósito:** contenido del paquete cerrado de pruebas + casos de prueba ejecutables (CT) + criterio de pase. El Ingeniero ejecuta los casos con el instructivo Capa B y declara pase/falla por caso.
> **Rector:** `Docs/Propuesta_Proyecto_Automatizacion_Remuneracion_UAESP.md` §9 Fase 3 it. 3.4, §10 (CA-4, CA-6) y §10.2.
> **Fecha de corte:** 2026-09-09 (HU-17, última historia 17/17).

---

## 1. Contenido del paquete

| Componente | Ubicación | Nota |
|---|---|---|
| Binario UI (Windows Forms) | `Remuneracion.WinForms/bin/Release/net10.0-windows/publish/` | Publicación framework-dependent (sin instalador; requiere runtime .NET 10) |
| Binario CLI | `Remuneracion.Cli/bin/Release/net10.0/publish/` | ídem |
| Código fuente + solución | Raíz del repo (`Remuneracion.WinForms/Remuneracion.WinForms.slnx`, 5 proyectos) | Build 0 warnings / 0 errores; tests 273/273 |
| Plantillas canónicas (solo lectura) | `Docs/Insumos/Remuneracion 202607-1 Total.xlsx`, `Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx` | SHA256 en el acta (§2 del instructivo); jamás mutar |
| Goldens de referencia (solo lectura) | `Docs/Insumos/Remuneracion 202607-1 Total.xlsx` (Q1) y `Remuneracion 202607-2 Total.xlsx` (Q2) | Oráculo de Capa B; nunca "ajustar" |
| Control Q2 (estructura, no oráculo) | `Docs/Insumos/REMUNERACION 2026072/Plantilla  _ Remuneracion 202607-2 Total.xlsx` | No se usa para comparar valores |
| Manual de usuario | `Docs/Manual-Usuario-Remuneracion-UAESP.md` | Qué hace la app, UI, CLI, códigos, catálogo, log, FAQ |
| Instructivo Capa B | `Docs/Instructivo-Capa-B.md` | Protocolo post-Excel con acta SHA y registro de evidencia |
| Este documento | `Docs/Paquete-Entrega-Pruebas.md` | Casos CT + criterio de pase |
| Registro de evidencia | `Docs/Instructivo-Capa-B.md` §10 | Tabla en blanco para una fila por caso |
| Logs | `remuneracion_log_*.txt` (rolling 30 días) | Evidencia CA-6 por RunId |

---

## 2. Casos de prueba

> Ejecute desde la raíz del repo (o con rutas absolutas). `--salida` es SIEMPRE una carpeta; el archivo se nombra solo (`Remuneración AAAAMM-# Total.xlsx`).

### CT-Q1-5A — Q1, 5 ASE, Capa B contra golden Q1

```bash
Remuneracion.Cli --periodo 2026071 --carpeta "Docs/Insumos/REMUNERACION 2026071" --plantilla "Docs/Insumos/Remuneracion 202607-1 Total.xlsx" --salida "Docs/Insumos/Salidas" --cinco-ase --sobrescribir
```

**Esperado:** exit **0** + línea `RESULTADO OK codigo=0 ... runId=<guid>` + archivo `Remuneración 202607-1 Total.xlsx` en `Docs/Insumos/Salidas`.
**Capa B:** bloques Q1.1–Q1.12 del instructivo vs golden Q1 (cadena D9:D109, DetRetri, REMUNERACION_*, BCE D/E, VALIDACION O=0/P=TRUE, INTERVENTORIA) — todo ±0.5 y fórmulas intactas.

### CT-Q1-1A — Q1, ASE único 3, Capa B contra golden Q1 (bloque ASE3)

```bash
Remuneracion.Cli --periodo 2026071 --carpeta "Docs/Insumos/REMUNERACION 2026071" --plantilla "Docs/Insumos/Remuneracion 202607-1 Total.xlsx" --salida "Docs/Insumos/Salidas" --ase 3
```

**Esperado:** exit **0** + `RESULTADO OK` + salida con el bloque ASE3 diligenciado.
**Capa B:** bloque ASE3 (D11/D30/D49/D68/D86/D106 y hojas R1/R2/R4/Recaudo del ASE 3) vs golden Q1 — ±0.5 y fórmulas intactas.

### CT-Q2-5A — Q2, 5 ASE, Capa B contra golden Q2

```bash
Remuneracion.Cli --periodo 2026072 --carpeta "Docs/Insumos/REMUNERACION 2026072" --plantilla "Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx" --salida "Docs/Insumos/Salidas" --cinco-ase --sobrescribir
```

**Esperado:** exit **0** + `RESULTADO OK` + archivo `Remuneración 202607-2 Total.xlsx`.
**Capa B:** bloques Q2.1–Q2.10 del instructivo vs golden Q2: AJUSTES D85:D89 (973693.46 / 216025.77 / 104231.83 / 35954.44 / 0), DetRetri-D enteros (D14 = 72441209168), D104:D109, INTERVENTORIA idéntica a Q1 — todo ±0.5 y fórmulas intactas.

### CT-NEG — Negativos (contrato de códigos)

| # | Comando (fragmento) | Esperado |
|---|---|---|
| N1 | `--periodo 2026071 --carpeta "C:\NoExiste" --plantilla <Q1> --salida <dir>` | exit **2** (`ERR-FUENTE-NO-ENCONTRADA`) |
| N2 | `--plantilla <misma-ruta-que-la-salida>` (salida == plantilla) | exit **2** (`ERR-PLANTILLA`) |
| N3 | Salida ya existe, **sin** `--sobrescribir` | exit **5** (`WARN-CANCELADO`) |
| N4 | `--periodo 2026` (período malformado) | exit **4** (uso; stderr con ayuda) |

Verifique además que la línea `RESULTADO ERROR codigo=<ERR-…> salida=<N> runId=<guid>` se imprime en N1–N3.

---

## 3. Criterio de pase

**El caso PASA** si y solo si:

1. **§10.2 del Rector (literal):** los valores del CONSOLIDADO_TOTAL RECAUDO generados automáticamente **coinciden exactamente** (o con diferencia **≤ ±0.5** por redondeo) con los valores del resultado manual conocido (golden) para el mismo período — verificado bloque a bloque en Capa B.
2. **CA-4:** las fórmulas de la plantilla **NO fueron alteradas** (spot-check de protegidas: D104:D109, AJUSTES-SF-T D47:D51, DetRetri/DetValiRetri D23:D28/D32:D36, REMUNERACION_*, BCE filas 9/11 + H, INTERVENTORIA SUM).
3. **CA-6:** el log registra **cada paso** de la ejecución, filtrable por RunId (inicio → fin → hitos por ASE/hoja).
4. Código de salida esperado del caso (0 en CT-Q1-5A / CT-Q1-1A / CT-Q2-5A; 2/5/4 en CT-NEG) y, en los positivos, hash de la plantilla intacto.

**Cualquier divergencia = FALLA** con registro de celda y valores (salida vs golden); nunca se "ajusta" el golden. Con todos los casos PASA, la entrega queda certificada y el proyecto cerrado (17/17 HU).

---

## 4. Evidencia de la fase de desarrollo (para contexto)

- Build: `dotnet build Remuneracion.WinForms/Remuneracion.WinForms.slnx` → **0 warnings / 0 errores**.
- Tests: `Remuneracion.IntegrationTests` → **273/273** (HU-01..HU-17, incluye paridad CLI↔UI).
- Harness de verificación (fuera de la solución): `Herramientas/VerificadorRecaudo` → **24/24**.
- Regresión Q1/Q2 intacta por construcción (cero cambios de cálculo/escritura/validación en HU-17).
