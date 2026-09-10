using Remuneracion.Core.Errors;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;

namespace Remuneracion.Core.Services;

/// <summary>
/// Validador de dominio de la remuneración.
/// HU-07: overload multi-ASE con matcheo ESTRICTO por <see cref="Ase.Id"/> (se elimina el
/// fallback a <see cref="Enumerable.FirstOrDefault"/>); el gate R1 compara la primera fila
/// Mes/Total (F25-equivalente por bloque) contra <see cref="ConsolidadoAse.Extemp"/> — nunca
/// <see cref="ConsolidadoAse.TotOpt"/> contra visibles R1.
/// </summary>
public sealed class ValidadorBasico : IValidador
{
    private const decimal Tolerancia = 0.5m;

    /// <summary>
    /// HU-14 (3.1): agrega un error con el prefijo <c>[ERR-VALIDACION]</c> del catálogo
    /// (todo error del validador es una validación de dominio que no cierra).
    /// </summary>
    private static void AgregarError(List<string> errores, string mensaje) =>
        errores.Add($"[{CodigoError.Validacion}] {mensaje}");

    public List<string> Validar(ResultadoRemuneracion resultado)
    {
        ArgumentNullException.ThrowIfNull(resultado);

        var errores = new List<string>();

        if (resultado.Consolidados.Count != 1)
        {
            AgregarError(errores, "Debe existir exactamente 1 consolidado para validar la hoja de salida en Fase 1.");
        }

        if (resultado.Consolidados.Count > 0 && resultado.Consolidados.Any(c => Math.Abs(c.AjustesSfT) > Tolerancia))
        {
            AgregarError(errores, $"AjustesSfT debe ser 0 en Fase 1, pero se encontró {resultado.Consolidados.First().AjustesSfT}.");
        }

        var consolidado = resultado.Consolidados.FirstOrDefault();
        if (consolidado is not null)
        {
            var diffGranTotal = Math.Abs(resultado.GranTotal - consolidado.TotalAse);
            if (diffGranTotal > Tolerancia)
            {
                AgregarError(errores, $"GranTotal no coincide con el único TotalAse: {resultado.GranTotal} vs {consolidado.TotalAse}. Diferencia={diffGranTotal}.");
            }
        }

        if (resultado.Consolidados.Count == 1)
        {
            var totalAse = resultado.Consolidados[0].TotalAse;
            var granTotal = resultado.GranTotal;
            if (Math.Abs(granTotal - totalAse) > Tolerancia)
            {
                AgregarError(errores, $"GranTotal ≠ único TotalAse ({granTotal} ≠ {totalAse}).");
            }
        }

        return errores;
    }

    public List<string> Validar(ResultadoRemuneracion resultado, WorkbookLeafInputs leaf)
    {
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leaf);

        var errores = Validar(resultado);

        // HU-07: matcheo estricto por Ase.Id — si falta el ASE, es error; nunca fallback.
        var consolidado = resultado.Consolidados.SingleOrDefault(c => c.Ase.Id == leaf.Ase.Id);
        if (consolidado is null)
        {
            AgregarError(errores, $"No se encontró el consolidado del ASE {leaf.Ase.Id} para comparar contra los inputs leaf.");
            return errores;
        }

        if (Math.Abs(consolidado.AjustesSfT) > Tolerancia)
        {
            AgregarError(errores, $"AjustesSfT debe ser 0: {consolidado.AjustesSfT}.");
        }

        ValidarGatesPorAse(errores, leaf, consolidado);

        return errores;
    }

    public List<string> Validar(ResultadoRemuneracion resultado, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leafs);

        var errores = new List<string>();

        // §2.5 regla 1: Consolidados.Count == leafs.Count; en modo período se exige 5.
        if (resultado.Consolidados.Count != leafs.Count)
        {
            AgregarError(errores, $"El modo período exige la misma cantidad de consolidados y leafs: {resultado.Consolidados.Count} consolidados vs {leafs.Count} leafs.");
        }

        if (leafs.Count != 5)
        {
            AgregarError(errores,
                leafs.Count == 1
                    ? "El modo período exige exactamente 5 ASE; se recibió 1 (¿usó el modo single-ASE?)."
                    : $"El modo período exige exactamente 5 ASE; se recibieron {leafs.Count}.");
        }

        // §2.5 regla 2: sin duplicados de Ase.Id.
        var duplicados = leafs
            .GroupBy(l => l.Ase.Id)
            .FirstOrDefault(g => g.Count() > 1);
        if (duplicados is not null)
        {
            AgregarError(errores, $"Se detectaron leafs con ASE duplicado: {duplicados.Key}.");
        }

        // §2.5 regla 3: AjustesSfT por quincena — Q1 exige 0 (intacto); Q2 exige el TotalAjustes
        // de la composición T0-0.3 ±0.5 con matcheo estricto por Ase.Id (D5). Si en Q2 un leaf no
        // trae AjustesSfT, es error que nombra el ASE (nunca 0 silencioso).
        var esQuincena2 = resultado.Periodo.NumeroQuincena == 2;
        foreach (var consolidado in resultado.Consolidados)
        {
            if (!esQuincena2)
            {
                if (Math.Abs(consolidado.AjustesSfT) > Tolerancia)
                {
                    AgregarError(errores, $"AjustesSfT debe ser 0 en Q1; ASE {consolidado.Ase.Id} tiene {consolidado.AjustesSfT}.");
                }

                continue;
            }

            var ajustes = leafs.SingleOrDefault(l => l.Ase.Id == consolidado.Ase.Id)?.AjustesSfT;
            if (ajustes is null)
            {
                AgregarError(errores, $"ASE {consolidado.Ase.Id}: en Q2 el leaf debe traer AjustesSfT (SALDOS POR NOTA + RETRIBUCION NEGATIVA) para validar el gate; no puede ser null.");
                continue;
            }

            var diferencia = Math.Abs(consolidado.AjustesSfT - ajustes.TotalAjustes);
            if (diferencia > Tolerancia)
            {
                AgregarError(errores, $"ASE {consolidado.Ase.Id}: AjustesSfT del consolidado ({consolidado.AjustesSfT}) no coincide con TotalAjustes de la fuente ({ajustes.TotalAjustes}). Diferencia={diferencia} > ±{Tolerancia}.");
            }
        }

        // §2.5 regla 4: gate leaf-vs-consolidado POR ASE con matcheo estricto (Single por Id).
        foreach (var leaf in leafs)
        {
            var consolidado = resultado.Consolidados.SingleOrDefault(c => c.Ase.Id == leaf.Ase.Id);
            if (consolidado is null)
            {
                AgregarError(errores, $"No se encontró el consolidado del ASE {leaf.Ase.Id} para el gate leaf-vs-consolidado.");
                continue;
            }

            ValidarGatesPorAse(errores, leaf, consolidado);

            // §2.5 regla 2.2 (HU-08): gate Σ empresas = visible de bloque por ASE y hoja
            // (R1/R2/R4), tolerancia ±0.5, ceros legítimos. Lista vacía = comportamiento HU-07.
            ValidarSigmaEmpresasPorAse(errores, leaf);
        }

        // HU-12 (2.6 ampliada, §2.5 regla 2): DetRetri-Q2 por ASE — SOLO en Q2 (Q1 sin cambios,
        // regresión ciega). Cada leaf debe traer DetRetriQ2 (error que nombra el ASE; nunca null)
        // y el Detalle debe coincidir con ROUND(D104:D108,0) vía DetRetriRounder (composición V0.4;
        // la comparación contra el caché golden vive en la Capa A, no en gates de fórmula).
        if (esQuincena2)
        {
            foreach (var leaf in leafs)
            {
                var detalle = leaf.DetRetriQ2;
                if (detalle is null)
                {
                    AgregarError(errores, $"ASE {leaf.Ase.Id}: en Q2 el leaf debe traer DetRetriQ2 (ROUND(D104:D108,0)) para validar el gate; no puede ser null.");
                    continue;
                }

                var redondeado = DetRetriRounder.Round(detalle.TotalD104);
                var diferencia = Math.Abs(detalle.Detalle - redondeado);
                if (diferencia > Tolerancia)
                {
                    AgregarError(errores, $"ASE {leaf.Ase.Id}: Detalle DetRetri ({detalle.Detalle}) no coincide con ROUND(D104:D108,0) ({redondeado}). Diferencia={diferencia} > ±{Tolerancia}.");
                }
            }
        }

        // §2.5 regla 5: GranTotal == Σ TotalAse (±0.5) + TotalAse aritmético por ASE.
        var sumaTotalAse = resultado.Consolidados.Sum(c => c.TotalAse);
        if (Math.Abs(resultado.GranTotal - sumaTotalAse) > Tolerancia)
        {
            AgregarError(errores, $"GranTotal ({resultado.GranTotal}) no coincide con Σ TotalAse ({sumaTotalAse}).");
        }

        foreach (var consolidado in resultado.Consolidados)
        {
            var totalAritmetico = consolidado.TotOpt
                + consolidado.R2TotalOportuno
                + consolidado.Extemp
                + consolidado.ReversionR4
                + consolidado.AjustesSfT;
            if (Math.Abs(consolidado.TotalAse - totalAritmetico) > Tolerancia)
            {
                AgregarError(errores, $"TotalAse del ASE {consolidado.Ase.Id} ({consolidado.TotalAse}) no coincide con su aritmética ({totalAritmetico}).");
            }
        }

        // HU-09 (2.3, §2.5 reglas 2-4): gates D5 del reporte por banco.
        ValidarReporteBanco(errores, leafs);

        // HU-10 (2.4, §2.5 reglas 2-4): gates D5 del balance de subsidios y contribuciones.
        ValidarBalanceSc(errores, leafs);

        return errores;
    }

    public List<string> Validar(
        ResultadoRemuneracion resultado,
        IReadOnlyList<WorkbookLeafInputs> leafs,
        IReadOnlyList<ValidacionCruzadaSnapshot> snapshots)
    {
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leafs);
        ArgumentNullException.ThrowIfNull(snapshots);

        // Todo HU-01..HU-12 intacto (gates existentes con matcheo estricto por Ase.Id).
        var errores = Validar(resultado, leafs);

        // §2.5 regla 2.7-0: la lista de snapshots debe tener la misma cantidad de ASE que los
        // consolidados/leafs. Snapshot ausente/vacío para un ASE = HU-12 puro para ese ASE.
        if (resultado.Consolidados.Count != snapshots.Count)
        {
            AgregarError(errores, $"El modo validaciones cruzadas exige la misma cantidad de consolidados y snapshots: {resultado.Consolidados.Count} consolidados vs {snapshots.Count} snapshots.");
        }

        if (snapshots.Count != 5)
        {
            AgregarError(errores, $"El modo validaciones cruzadas exige exactamente 5 snapshots (uno por ASE); se recibieron {snapshots.Count}.");
        }

        var snapshotsOrdenados = snapshots.OrderBy(s => s.Ase.Id).ToList();
        foreach (var snapshot in snapshotsOrdenados)
        {
            ValidarGatesValidacionesCruzadasPorAse(errores, snapshot, resultado);
        }

        // HU-14 (S-2, D8): VALIDACION_TOTAL con gate ÚNICO (no ×5) + igualdad entre snapshots.
        // El reader lee O9/P9 una sola vez; el gate evalúa el primer snapshot y la igualdad
        // detecta una divergencia del lector (fail-fast con 1 error, no 5).
        if (snapshotsOrdenados.Count > 0)
        {
            var baseTotal = snapshotsOrdenados[0];
            if (Math.Abs(baseTotal.ValidacionTotal) > Tolerancia)
            {
                AgregarError(errores, $"ASE {baseTotal.Ase.Id}: VALIDACION_TOTAL (O9) no cierra: {baseTotal.ValidacionTotal} (debe ser 0 ±{Tolerancia}).");
            }

            if (!baseTotal.ValidacionTotalOkP)
            {
                AgregarError(errores, $"ASE {baseTotal.Ase.Id}: VALIDACION_TOTAL P9 (INT(O9)=0) es falso en el workbook.");
            }

            foreach (var snapshot in snapshotsOrdenados.Skip(1))
            {
                if (snapshot.ValidacionTotal != baseTotal.ValidacionTotal
                    || snapshot.ValidacionTotalOkP != baseTotal.ValidacionTotalOkP)
                {
                    AgregarError(errores, $"VALIDACION_TOTAL: el snapshot del ASE {snapshot.Ase.Id} (O9={snapshot.ValidacionTotal}, P9={snapshot.ValidacionTotalOkP}) diverge del ASE {baseTotal.Ase.Id}; el lector-oráculo debe leer O9/P9 una sola vez (S-2).");
                    break;
                }
            }
        }

        return errores;
    }

    /// <summary>
    /// HU-13 (2.7, §2.5): gates aditivos por ASE contra el snapshot-oráculo (D1/D3). El validador
    /// NO abre .xlsx: solo compara números. Matcheo estricto por <see cref="Ase.Id"/> (Single,
    /// nunca fallback). Semántica congelada por T0 en ambos canónicos (Plan 13 §4 Fase 0).
    ///
    /// HU-14 (W-2, veredicto documentado — sin mover): este gate de validaciones cruzadas vive
    /// AQUÍ (ValidadorBasico) y NO en <c>WorkbookLeafCoherence</c> por diseño: todos los gates
    /// cruzados HU-08..HU-12 (<c>ValidarSigmaEmpresasPorAse</c>, <c>ValidarReporteBanco</c>,
    /// <c>ValidarBalanceSc</c>) viven en el validador con AGREGACIÓN de errores por lista;
    /// moverlos a <c>WorkbookLeafCoherence</c> (throw al primer fallo) cambiaría la semántica
    /// de agregación. Ver Plan 13 §2.8/§4-3.1 enmendado por Plan 14 §9.7.
    ///
    /// HU-14 (W-1): además de O/P por empresa y DetValiRetri, se gatean los sub-bloques
    /// booleanos de <c>VALIDACION_TOTAL</c> (C15/D25/O25/D34/F34) == true exacto por ASE.
    /// HU-14 (S-2): el gate numérico/booleano de VALIDACION_TOTAL (O9/P9) NO se evalúa aquí
    /// por ASE (sería ×5); se evalúa UNA vez en <c>Validar(resultado, leafs, snapshots)</c>.
    /// </summary>
    private static void ValidarGatesValidacionesCruzadasPorAse(
        List<string> errores,
        ValidacionCruzadaSnapshot snapshot,
        ResultadoRemuneracion resultado)
    {
        ArgumentNullException.ThrowIfNull(snapshot);

        var consolidado = resultado.Consolidados.SingleOrDefault(c => c.Ase.Id == snapshot.Ase.Id);
        if (consolidado is null)
        {
            AgregarError(errores, $"No se encontró el consolidado del ASE {snapshot.Ase.Id} para el gate de validaciones cruzadas.");
            return;
        }

        // (i) Por empresa (VALIDACION_*): fila 2+aseId de cada hoja. O = H−N == 0 ±0.5 y
        //     P (INT(O)=0) == true exacto (falla aunque O esté en tolerancia). Error ASE+empresa.
        foreach (var empresa in snapshot.PorEmpresa)
        {
            if (empresa.AseId != snapshot.Ase.Id)
            {
                AgregarError(errores, $"ASE {snapshot.Ase.Id}: el snapshot por empresa trae la fila del ASE {empresa.AseId} (empresa {empresa.Empresa}); matcheo estricto violado.");
                continue;
            }

            var diferenciaO = Math.Abs(empresa.DiferenciaO);
            if (diferenciaO > Tolerancia)
            {
                AgregarError(errores, $"ASE {snapshot.Ase.Id} · {empresa.Empresa}: la validación Recaudo vs REMUNERACION (O) no cierra: O={empresa.DiferenciaO} (debe ser 0 ±{Tolerancia}).");
            }

            if (!empresa.VerificacionP)
            {
                AgregarError(errores, $"ASE {snapshot.Ase.Id} · {empresa.Empresa}: la verificación P (INT(O)=0) es falsa en el workbook aunque O esté en tolerancia.");
            }
        }

        // (ii) DetValiRetri: D16..D20 (diferencias por ASE) == 0 ±0.5; D24..D28 == true exacto;
        //     D29 (total) == true exacto. D21 (fila Total) y D9:D14/J9:J14 ≠ ROUND quedan
        //     EXCLUIDOS (D6: divergencia documentada en ambos canónicos).
        if (snapshot.DetValiRetri is not null)
        {
            foreach (var celda in snapshot.DetValiRetri.DiferenciasAse)
            {
                var diferencia = Math.Abs(celda.Valor);
                if (diferencia > Tolerancia)
                {
                    AgregarError(errores, $"ASE {snapshot.Ase.Id}: DetValiRetri {celda.Celda} no cierra: {celda.Valor} (debe ser 0 ±{Tolerancia}).");
                }
            }

            foreach (var celda in snapshot.DetValiRetri.VerificacionesAse)
            {
                if (!celda.Verificacion)
                {
                    AgregarError(errores, $"ASE {snapshot.Ase.Id}: DetValiRetri {celda.Celda} es falso (composición no verificada).");
                }
            }

            if (!snapshot.DetValiRetri.VerificacionTotalD29)
            {
                AgregarError(errores, $"ASE {snapshot.Ase.Id}: DetValiRetri D29 (verificación TOTAL) es falso.");
            }
        }

        // (iii) HU-14 (W-1): sub-bloques booleanos de VALIDACION_TOTAL (C15/D25/O25/D34/F34)
        //     == true exacto por ASE (amparo T0-0.4 HU-13: TRUE en ambos goldens). El gate
        //     numérico/booleano TOTAL (O9/P9) es ÚNICO (S-2) y vive en Validar(...snapshots).
        foreach (var (celda, valor) in snapshot.SubBloquesValidacionTotal.OrderBy(kv => kv.Key))
        {
            if (!valor)
            {
                AgregarError(errores, $"ASE {snapshot.Ase.Id} · VALIDACION_TOTAL · {celda}: sub-bloque booleano es falso en el workbook (debe ser TRUE exacto; W-1).");
            }
        }
    }

    /// <summary>
    /// HU-10 (2.4, §2.5): gates D5 del balance SC.
    /// (i) Por ASE: Total BSC (Contribucion+Subsidio) == Total General fuente (col G) ±0.5;
    ///     fail-fast nombra el ASE.
    /// (ii) Σ fila 11 en dominio: Σ Total BSC a través de los 5 ASE == Σ TotalFuente ±0.5
    ///     (las celdas D11/E11/F11 del template son fórmulas protegidas; la aritmética de
    ///     dominio verifica que la suma cierra).
    /// (iii) H≈F solo cuando el desenlace T0 es D2(a) (<see cref="BalanceScAseInputs.Sistema"/>
    ///     no nulo): redondeo a entero (AwayFromZero). En Q1 el veredicto es D2(b) (H es
    ///     fórmula → Sistema null) y H≈F se verifica contra el caché golden en Capa A (G7).
    /// Todos los leafs con <c>BalanceSc == null</c> = comportamiento HU-09 puro (sin gates 2.4).
    /// </summary>
    private static void ValidarBalanceSc(List<string> errores, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        var conBalance = leafs.Where(l => l.BalanceSc is not null).ToList();
        if (conBalance.Count == 0)
        {
            return; // HU-09 puro: sin gates 2.4
        }

        foreach (var leaf in conBalance)
        {
            var balance = leaf.BalanceSc!;
            var bloque = balance.Ases.SingleOrDefault(b => b.Ase.Id == leaf.Ase.Id);
            if (bloque is null)
            {
                AgregarError(errores, $"El leaf del ASE {leaf.Ase.Id} no trae su fila de balance SC para validar.");
                continue;
            }

            // Gate (i): BCE = fuente ±0.5 (TotalBsc vs TotalFuente; col G del Total General).
            var diferencia = Math.Abs(bloque.TotalBsc - bloque.TotalFuente);
            if (diferencia > Tolerancia)
            {
                AgregarError(errores, $"ASE {leaf.Ase.Id}: Total BSC ({bloque.TotalBsc}) != Total General fuente ({bloque.TotalFuente}). Diferencia={diferencia} > ±{Tolerancia}.");
            }

            // Gate (iii): H≈F con redondeo a entero SOLO en desenlace D2(a); en D2(b) (Q1)
            // Sistema es null y H≈F se verifica en Capa A contra el caché golden.
            if (bloque.Sistema is not null)
            {
                var redondeado = Math.Round(bloque.TotalBsc, MidpointRounding.AwayFromZero);
                if (Math.Abs(bloque.Sistema.Value - redondeado) > Tolerancia)
                {
                    AgregarError(errores, $"ASE {leaf.Ase.Id}: H (SISTEMA) {bloque.Sistema.Value} no coincide con ROUND(F,0)={redondeado}.");
                }
            }
        }

        // Gate (ii): sumas fila 11 en dominio — Σ Total BSC == Σ TotalFuente ±0.5.
        var sumaBsc = conBalance.Sum(l => l.BalanceSc!.Ases.Sum(b => b.TotalBsc));
        var sumaFuente = conBalance.Sum(l => l.BalanceSc!.Ases.Sum(b => b.TotalFuente));
        if (Math.Abs(sumaBsc - sumaFuente) > Tolerancia)
        {
            AgregarError(errores, $"Σ Total BSC ({sumaBsc}) != Σ Total General fuente ({sumaFuente}). Diferencia={Math.Abs(sumaBsc - sumaFuente)} > ±{Tolerancia}.");
        }
    }

    /// <summary>
    /// HU-09 (2.3, §2.5): gates D5 del reporte por banco.
    /// (i) Por ASE y empresa: bloque (Σ conceptos) == resumen fuente (<see cref="ReporteBancoEmpresaInputs.TotalFuente"/>)
    ///     ±0.5; fail-fast nombra ASE y empresa-columna.
    /// (ii) Consolidado 1–7 == Σ bloques por empresa ±0.5: si el dominio trae
    ///     <see cref="ReporteBancoInputs.Consolidado"/> se compara y falla si difiere; en el
    ///     desenlace T0 D2(b) (filas 1–7 son fórmulas) el campo es null y la verificación Σ
    ///     contra el caché golden vive en la Capa A (G3: "solo se verifica Σ").
    /// (iii) C59 (<see cref="ReporteBancoInputs.Quincena"/>) == <see cref="Periodo.NumeroQuincena"/>.
    /// Todos los leafs con <c>ReporteBanco == null</c> = comportamiento HU-08 puro (sin gates 2.3).
    /// </summary>
    private static void ValidarReporteBanco(List<string> errores, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        var conBanco = leafs.Where(l => l.ReporteBanco is not null).ToList();
        if (conBanco.Count == 0)
        {
            return; // HU-08 puro: sin gates 2.3
        }

        foreach (var leaf in conBanco)
        {
            var banco = leaf.ReporteBanco!;

            // Gate (iii): C59 == quincena (dominio, nunca fuente).
            if (banco.Quincena != leaf.Periodo.NumeroQuincena)
            {
                AgregarError(errores, $"ASE {leaf.Ase.Id}: C59 del reporte por banco ({banco.Quincena}) no coincide con la quincena del período ({leaf.Periodo.NumeroQuincena}).");
            }

            // Gate (i): bloque = resumen fuente por empresa ±0.5 (fail-fast nombra ASE + empresa).
            foreach (var bloque in banco.Ases)
            {
                foreach (var empresa in bloque.Empresas)
                {
                    var diferencia = Math.Abs(empresa.Total - empresa.TotalFuente);
                    if (diferencia > Tolerancia)
                    {
                        AgregarError(errores, $"ASE {leaf.Ase.Id} · {empresa.Empresa}: bloque banco ({empresa.Total}) != resumen fuente ({empresa.TotalFuente}). Diferencia={diferencia} > ±{Tolerancia}.");
                    }
                }
            }
        }

        // Gate (ii): consolidado 1–7 == Σ bloques por empresa ±0.5. En D2(b) el Consolidado
        // del dominio es null (filas 1–7 son fórmulas): la Σ se verifica contra el caché
        // golden en Capa A y se expone en el log/UI, nunca como error aquí.
        var consolidado = conBanco.Select(l => l.ReporteBanco!.Consolidado).FirstOrDefault(c => c is not null);
        if (consolidado is null)
        {
            return;
        }

        var sumaPorEmpresa = conBanco
            .SelectMany(l => l.ReporteBanco!.Ases)
            .SelectMany(b => b.Empresas)
            .GroupBy(e => e.Empresa, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Sum(e => e.Total), StringComparer.OrdinalIgnoreCase);

        foreach (var (empresa, esperado) in consolidado)
        {
            if (!sumaPorEmpresa.TryGetValue(empresa, out var suma))
            {
                AgregarError(errores, $"Consolidado banco declara la empresa '{empresa}' pero no hay bloques con esa empresa en los leafs.");
                continue;
            }

            if (Math.Abs(suma - esperado) > Tolerancia)
            {
                AgregarError(errores, $"Consolidado 1–7 de '{empresa}' ({esperado}) no coincide con Σ bloques por empresa ({suma}).");
            }
        }
    }

    private static void ValidarGatesPorAse(
        List<string> errores,
        WorkbookLeafInputs leaf,
        ConsolidadoAse consolidado)
    {
        // Gate R1: primera fila Mes/Total col F (F25-equivalente por bloque) vs Extemp.
        // Prohibido comparar TotOpt HU-02 contra visibles R1.
        if (Math.Abs(leaf.R1.F25 - consolidado.Extemp) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: la hoja R1 F25-equivalente ({leaf.R1.F25}) no coincide con Extemp del consolidado ({consolidado.Extemp}).");
        }

        // Gate R2: visible del bloque vs R2TotalOportuno.
        if (Math.Abs(leaf.R2.TotalOportunoEsperado - consolidado.R2TotalOportuno) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: la hoja R2 TotalOportunoEsperado ({leaf.R2.TotalOportunoEsperado}) no coincide con R2TotalOportuno del consolidado ({consolidado.R2TotalOportuno}).");
        }

        // Gate R4: visible del bloque vs ReversionR4.
        if (Math.Abs(leaf.R4.TotalReversionEsperada - consolidado.ReversionR4) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: la hoja R4 TotalReversionEsperada ({leaf.R4.TotalReversionEsperada}) no coincide con ReversionR4 del consolidado ({consolidado.ReversionR4}).");
        }
    }

    /// <summary>
    /// HU-08 (§2.5): Σ visibles-empresa == visible de bloque por hoja (R1/R2/R4) ±0.5.
    /// Ceros legítimos: EAAB-CL todo 0 en Q1 es válido (V8). El error nombra ASE + empresa.
    /// Lista vacía = HU-07 puro.
    /// </summary>
    private static void ValidarSigmaEmpresasPorAse(List<string> errores, WorkbookLeafInputs leaf)
    {
        if (leaf.Conciliacion.Count == 0)
        {
            return;
        }

        var sumaR1 = leaf.Conciliacion.Sum(c => c.VisibleR1);
        if (Math.Abs(sumaR1 - leaf.R1.TotalOportunoEsperadoPorAse) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: Σ visibles de empresas en R1 ({sumaR1}) no coincide con el visible de bloque ({leaf.R1.TotalOportunoEsperadoPorAse}). Detalle: {DetalleEmpresas(leaf, c => c.VisibleR1)}");
        }

        var sumaR2 = leaf.Conciliacion.Sum(c => c.VisibleR2);
        if (Math.Abs(sumaR2 - leaf.R2.TotalOportunoEsperado) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: Σ visibles de empresas en R2 ({sumaR2}) no coincide con el visible de bloque ({leaf.R2.TotalOportunoEsperado}). Detalle: {DetalleEmpresas(leaf, c => c.VisibleR2)}");
        }

        var sumaR4 = leaf.Conciliacion.Sum(c => c.VisibleR4);
        if (Math.Abs(sumaR4 - leaf.R4.TotalReversionEsperada) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: Σ visibles de empresas en R4 ({sumaR4}) no coincide con el visible de bloque ({leaf.R4.TotalReversionEsperada}). Detalle: {DetalleEmpresas(leaf, c => c.VisibleR4)}");
        }
    }

    private static string DetalleEmpresas(WorkbookLeafInputs leaf, Func<ConciliacionEmpresaInputs, decimal> selector) =>
        string.Join(", ", leaf.Conciliacion.Select(c => $"{c.Empresa.Nombre}={selector(c)}"));
}
