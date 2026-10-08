using Remuneracion.Core.Models;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 30 (T1, R-C-1/S5): test parametrizado del helper Core <see cref="NombresHojaPeriodo"/>.
/// El nombre de hoja se COMPONE del dominio (<c>CodigoCompleto</c>); para los períodos canónicos
/// (2026071/2026072) resuelve byte-idéntico a los literales previos y el caso futuro (2026091)
/// demuestra que un período nuevo no requiere cambios de producción (OCP).
/// </summary>
public sealed class NombresHojaPeriodoTests
{
    [Theory]
    [InlineData("2026071", "DetRetri2026071", "DetValiRetri2026071", "Informe AFaseo Recaudo 202607-1")]
    [InlineData("2026072", "DetRetri2026072", "DetValiRetri2026072", "Informe AFaseo Recaudo 202607-2")]
    [InlineData("2026082", "DetRetri2026082", "DetValiRetri2026082", "Informe AFaseo Recaudo 202608-2")]
    [InlineData("2026091", "DetRetri2026091", "DetValiRetri2026091", "Informe AFaseo Recaudo 202609-1")]
    public void ComponeNombresDeHojaDesdeElDominio(
        string codigoCompleto, string detRetri, string detValiRetri, string informe)
    {
        var periodo = Periodo.Parse(codigoCompleto);

        // API por código completo (AAAAMMQ).
        Assert.Equal(detRetri, NombresHojaPeriodo.DetRetri(codigoCompleto));
        Assert.Equal(detValiRetri, NombresHojaPeriodo.DetValiRetri(codigoCompleto));
        Assert.Equal(informe, NombresHojaPeriodo.InformeAFaseo(periodo.CodigoAAAAMM, periodo.NumeroQuincena));

        // Sobrecargas por dominio → mismo resultado.
        Assert.Equal(detRetri, NombresHojaPeriodo.DetRetri(periodo));
        Assert.Equal(detValiRetri, NombresHojaPeriodo.DetValiRetri(periodo));
        Assert.Equal(informe, NombresHojaPeriodo.InformeAFaseo(periodo));
    }
}
