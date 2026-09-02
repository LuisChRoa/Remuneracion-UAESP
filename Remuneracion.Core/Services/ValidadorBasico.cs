using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// Validador básico de dominio para la Fase 1 (1 ASE, quincena 1).
/// </summary>
public sealed class ValidadorBasico : IValidador
{
    private const decimal Tolerancia = 0.5m;

    public List<string> Validar(ResultadoRemuneracion resultado)
    {
        ArgumentNullException.ThrowIfNull(resultado);

        var errores = new List<string>();

        if (resultado.Consolidados.Count != 1)
        {
            errores.Add("Debe existir exactamente 1 consolidado para validar la hoja de salida en Fase 1.");
        }

        if (resultado.Consolidados.Count > 0 && resultado.Consolidados.Any(c => Math.Abs(c.AjustesSfT) > Tolerancia))
        {
            errores.Add($"AjustesSfT debe ser 0 en Fase 1, pero se encontró {resultado.Consolidados.First().AjustesSfT}.");
        }

        var consolidado = resultado.Consolidados.FirstOrDefault();
        if (consolidado is not null)
        {
            var diffGranTotal = Math.Abs(resultado.GranTotal - consolidado.TotalAse);
            if (diffGranTotal > Tolerancia)
            {
                errores.Add($"GranTotal no coincide con el único TotalAse: {resultado.GranTotal} vs {consolidado.TotalAse}. Diferencia={diffGranTotal}.");
            }
        }

        if (resultado.Consolidados.Count == 1)
        {
            var totalAse = resultado.Consolidados[0].TotalAse;
            var granTotal = resultado.GranTotal;
            if (Math.Abs(granTotal - totalAse) > Tolerancia)
            {
                errores.Add($"GranTotal ≠ único TotalAse ({granTotal} ≠ {totalAse}).");
            }
        }

        return errores;
    }

    public List<string> Validar(ResultadoRemuneracion resultado, WorkbookLeafInputs leaf)
    {
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leaf);

        var errores = Validar(resultado);

        var consolidado = resultado.Consolidados.FirstOrDefault(c => c.Ase.Id == leaf.Ase.Id)
            ?? resultado.Consolidados.FirstOrDefault();

        if (consolidado is null)
        {
            errores.Add("No se encontró un consolidado válido para comparar contra los inputs leaf.");
            return errores;
        }

        if (Math.Abs(consolidado.AjustesSfT) > Tolerancia)
        {
            errores.Add($"AjustesSfT debe ser 0: {consolidado.AjustesSfT}.");
        }

        if (Math.Abs(leaf.R1.F25 - consolidado.Extemp) > Tolerancia)
        {
            errores.Add($"La hoja R1 F25 ({leaf.R1.F25}) no coincide con Extemp del consolidado ({consolidado.Extemp}).");
        }

        if (Math.Abs(leaf.R2.TotalOportunoEsperado - consolidado.R2TotalOportuno) > Tolerancia)
        {
            errores.Add($"La hoja R2 TotalOportunoEsperado ({leaf.R2.TotalOportunoEsperado}) no coincide con R2TotalOportuno del consolidado ({consolidado.R2TotalOportuno}).");
        }

        if (Math.Abs(leaf.R4.TotalReversionEsperada - consolidado.ReversionR4) > Tolerancia)
        {
            errores.Add($"La hoja R4 TotalReversionEsperada ({leaf.R4.TotalReversionEsperada}) no coincide con ReversionR4 del consolidado ({consolidado.ReversionR4}).");
        }

        return errores;
    }
}
