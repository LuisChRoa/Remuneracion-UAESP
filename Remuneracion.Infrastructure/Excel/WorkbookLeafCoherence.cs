using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

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
    }
}