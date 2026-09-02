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

        var consolidado = resultado.Consolidados.FirstOrDefault(c => c.Ase.Id == leaf.Ase.Id)
            ?? resultado.Consolidados.FirstOrDefault()
            ?? throw new CalculoInvalidoException(
                "No hay consolidado en ResultadoRemuneracion para validar coherencia leaf vs agregado antes de escribir.");

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

    internal static void AsegurarDentroDeTolerancia(string etiqueta, decimal leaf, decimal agregado)
    {
        var diferencia = Math.Abs(leaf - agregado);
        if (diferencia > Tolerancia)
        {
            throw new CalculoInvalidoException(
                $"La coherencia '{etiqueta}' no se cumple: leaf={leaf} vs agregado={agregado}. Diferencia={diferencia} > ±{Tolerancia}.");
        }
    }
}
