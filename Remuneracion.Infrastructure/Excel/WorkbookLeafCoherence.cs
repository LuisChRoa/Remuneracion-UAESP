using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Gate de coherencia leaf → agregado antes de certificar escritura.
/// R2 y R4 se comparan contra los agregados HU-02/HU-03.
/// R1 no compara <see cref="WorkbookLeafInputsR1.TotalOportunoEsperado"/> contra
/// <see cref="RecaudoComponenteR1.TotalOportuno"/>: en el workbook real F46 = F25+F41-L25
/// (TOT_OPT visible del consolidado) y el agregado HU-02 es la fila Componente/Total, otra cantidad.
/// Sí se compara F25 contra Extemporáneo HU-02 (primera fila Mes/Total col F).
///
/// HU-07: matcheo ESTRICTO por <see cref="Ase.Id"/> — se elimina el fallback a
/// <see cref="Enumerable.FirstOrDefault"/> (plan G5/D4).
/// </summary>
internal static class WorkbookLeafCoherence
{
    internal const decimal Tolerancia = 0.5m;

    internal static void ValidarContraFuentes(
        WorkbookLeafInputs leaf,
        RecaudoComponenteR1 r1,
        SaldosFavorR2 r2,
        ReversionR4 r4)
    {
        ArgumentNullException.ThrowIfNull(leaf);
        ArgumentNullException.ThrowIfNull(r1);
        ArgumentNullException.ThrowIfNull(r2);
        ArgumentNullException.ThrowIfNull(r4);

        AsegurarDentroDeTolerancia(
            "R1.F25 vs RecaudoComponenteR1.Extemporaneo",
            leaf.R1.F25,
            r1.Extemporaneo);
        AsegurarDentroDeTolerancia(
            "R2.TotalOportunoEsperado vs SaldosFavorR2.TotalOportuno",
            leaf.R2.TotalOportunoEsperado,
            r2.TotalOportuno);
        AsegurarDentroDeTolerancia(
            "R4.TotalReversionEsperada vs ReversionR4.TotalReversiones",
            leaf.R4.TotalReversionEsperada,
            r4.TotalReversiones);
    }

    internal static void ValidarContraResultado(WorkbookLeafInputs leaf, ResultadoRemuneracion resultado)
    {
        ArgumentNullException.ThrowIfNull(leaf);
        ArgumentNullException.ThrowIfNull(resultado);

        // HU-07: matcheo estricto — nunca fallback a otro consolidado.
        var consolidado = resultado.Consolidados.SingleOrDefault(c => c.Ase.Id == leaf.Ase.Id)
            ?? throw new CalculoInvalidoException(
                $"No hay consolidado del ASE {leaf.Ase.Id} en ResultadoRemuneracion para validar coherencia leaf vs agregado antes de escribir.");

        ValidarContraConsolidado(leaf, consolidado);
    }

    /// <summary>
    /// Gate multi-ASE: valida cada leaf contra su consolidado con matcheo estricto por Ase.Id.
    /// </summary>
    internal static void ValidarContraResultadoMultiAse(
        IReadOnlyList<WorkbookLeafInputs> leafs,
        ResultadoRemuneracion resultado)
    {
        ArgumentNullException.ThrowIfNull(leafs);
        ArgumentNullException.ThrowIfNull(resultado);

        if (leafs.Count == 0)
        {
            throw new CalculoInvalidoException("La lista de leafs del período está vacía; no hay nada que certificar.");
        }

        foreach (var leaf in leafs)
        {
            ValidarContraResultado(leaf, resultado);
        }
    }

    /// <summary>
    /// HU-08 (2.2): gate Σ empresas = visible de bloque por ASE y hoja (R1/R2/R4), tolerancia
    /// ±0.5, ceros legítimos (EAAB-CL todo 0 en Q1 es válido). Lista vacía = HU-07 puro.
    /// </summary>
    internal static void ValidarSigmaEmpresas(
        IReadOnlyList<ConciliacionEmpresaInputs> conciliacion,
        WorkbookLeafInputs leaf)
    {
        ArgumentNullException.ThrowIfNull(conciliacion);
        ArgumentNullException.ThrowIfNull(leaf);

        if (conciliacion.Count == 0)
        {
            return;
        }

        var sumaR1 = conciliacion.Sum(c => c.VisibleR1);
        var sumaR2 = conciliacion.Sum(c => c.VisibleR2);
        var sumaR4 = conciliacion.Sum(c => c.VisibleR4);

        AsegurarDentroDeTolerancia(
            $"ASE {leaf.Ase.Id} R1: Σ empresas vs visible de bloque (TOT_OPT)",
            sumaR1,
            leaf.R1.TotalOportunoEsperadoPorAse);
        AsegurarDentroDeTolerancia(
            $"ASE {leaf.Ase.Id} R2: Σ empresas vs visible de bloque",
            sumaR2,
            leaf.R2.TotalOportunoEsperado);
        AsegurarDentroDeTolerancia(
            $"ASE {leaf.Ase.Id} R4: Σ empresas vs visible de bloque",
            sumaR4,
            leaf.R4.TotalReversionEsperada);
    }

    internal static void AsegurarDentroDeTolerancia(string etiqueta, decimal leaf, decimal agregado)
    {
        var diferencia = Math.Abs(leaf - agregado);
        if (diferencia > Tolerancia)
        {
            throw new CalculoInvalidoException(
                $"La coherencia '{etiqueta}' no se cumple: leaf={leaf} vs agregado={agregado}. Diferencia={diferencia} > ±{Tolerancia}.");
        }
    }

    /// <summary>
    /// HU-09 (2.3, §2.5 gate D5-i): verifica que el bloque banco del ASE sea EXACTO contra el
    /// "Resumen Recaudo Aplicado Por Servicio" de la fuente (Σ conceptos == Total de la fila
    /// "Total" de la fuente, ±0.5). Fail-fast nombra ASE y empresa-columna. Es la coherencia
    /// "bloque template == resumen fuente" antes de escribir.
    /// </summary>
    internal static void ValidarContraFuentesBanco(
        ReporteBancoAseInputs bloque,
        IReadOnlyList<ReporteBancoEmpresaInputs> empresas)
    {
        ArgumentNullException.ThrowIfNull(bloque);
        ArgumentNullException.ThrowIfNull(empresas);

        foreach (var empresa in empresas)
        {
            AsegurarDentroDeTolerancia(
                $"ASE {bloque.Ase.Id} · {empresa.Empresa}: bloque banco vs resumen fuente (fila Total)",
                empresa.Total,
                empresa.TotalFuente);
        }
    }

    /// <summary>
    /// HU-09 (2.3): Σ consolidado 1–7 por empresa a través de los 5 leafs (aritmética de
    /// dominio). Se usa en la Capa A y en el log/UI para verificar que las filas 1–7 (fórmulas
    /// protegidas, D2(b)) computan Σ bloques por empresa.
    /// </summary>
    internal static IReadOnlyDictionary<string, decimal> SumaConsolidadoPorEmpresa(
        IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        ArgumentNullException.ThrowIfNull(leafs);

        return leafs
            .Where(l => l.ReporteBanco is not null)
            .SelectMany(l => l.ReporteBanco!.Ases)
            .SelectMany(b => b.Empresas)
            .GroupBy(e => e.Empresa, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Sum(e => e.Total), StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// HU-10 (2.4, §2.5 gate D5-i): verifica que la fila del balance del ASE sea EXACTA contra
    /// la fila "Total General" de la fuente (TotalBsc = Contribucion + Subsidio == TotalFuente,
    /// ±0.5). Fail-fast nombra el ASE. Es la coherencia "BCE = fuente" antes de escribir.
    /// </summary>
    internal static void ValidarContraFuentesBalanceSc(BalanceScAseInputs bloque)
    {
        ArgumentNullException.ThrowIfNull(bloque);

        AsegurarDentroDeTolerancia(
            $"ASE {bloque.Ase.Id}: Total BSC (Contribucion+Subsidio) vs Total General fuente (col G)",
            bloque.TotalBsc,
            bloque.TotalFuente);
    }

    private static void ValidarContraConsolidado(WorkbookLeafInputs leaf, ConsolidadoAse consolidado)
    {
        AsegurarDentroDeTolerancia(
            "R1.F25 vs ConsolidadoAse.Extemp",
            leaf.R1.F25,
            consolidado.Extemp);
        AsegurarDentroDeTolerancia(
            "R2.TotalOportunoEsperado vs ConsolidadoAse.R2TotalOportuno",
            leaf.R2.TotalOportunoEsperado,
            consolidado.R2TotalOportuno);
        AsegurarDentroDeTolerancia(
            "R4.TotalReversionEsperada vs ConsolidadoAse.ReversionR4",
            leaf.R4.TotalReversionEsperada,
            consolidado.ReversionR4);

        // HU-11 (2.5, §2.5 regla 6): coherencia leaf ajustes vs consolidado con matcheo estricto
        // por Ase.Id. AjustesSfT == null = comportamiento HU-10 puro (Q1) — sin gate.
        if (leaf.AjustesSfT is not null)
        {
            AsegurarDentroDeTolerancia(
                $"ASE {leaf.Ase.Id} AJUSTES-SF-T: TotalAjustes (composición T0-0.3) vs ConsolidadoAse.AjustesSfT",
                leaf.AjustesSfT.TotalAjustes,
                consolidado.AjustesSfT);
        }

        // HU-12 (2.6 ampliada, §2.5 regla 2): coherencia DetRetri-Q2 con matcheo estricto por
        // Ase.Id. DetRetriQ2 == null = comportamiento HU-11 puro (Q1) — sin gate. La composición
        // V0.4 es CONGELADA (Detalle = ROUND(D104:D108,0) vía DetRetriRounder); el writer falla
        // si el leaf declara un Detalle incoherente con su TotalD104.
        //
        // HU-13 (2.7, hallazgo T0): NO se agrega coherencia TotalD104 vs ConsolidadoAse.TotalAse —
        // en Q2 son cantidades distintas (TotalD104 replica CONSOLIDADO D104:D108 = RECAUDO TOTAL
        // del ASE; ConsolidadoAse.TotalAse es la REMUNERACIÓN total). Prohibido cruzarlas (A5).
        if (leaf.DetRetriQ2 is not null)
        {
            AsegurarDentroDeTolerancia(
                $"ASE {leaf.Ase.Id} DetRetri-Q2: Detalle vs ROUND(D104:D108,0)",
                leaf.DetRetriQ2.Detalle,
                DetRetriRounder.Round(leaf.DetRetriQ2.TotalD104));
        }
    }
}
