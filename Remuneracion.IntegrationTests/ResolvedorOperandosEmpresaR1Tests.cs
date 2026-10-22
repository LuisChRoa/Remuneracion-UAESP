using Remuneracion.Core.Models;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 53 (H1.1, D1/D4/D5): unit tests del resolvedor PURO de operandos-empresa R1 por sección
/// declarada (<see cref="ResolvedorOperandosEmpresaR1"/>). Sintéticos: la regla se prueba con
/// listas de filas ya tipadas (roles, ausencias → 0, empate → señal de STOP, única → Main).
/// La evidencia con fuentes reales vive en <see cref="ContratoSeccionPorOperandoR1Tests"/> y en
/// <c>Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md</c>.
/// </summary>
public sealed class ResolvedorOperandosEmpresaR1Tests
{
    private static readonly SeccionFronteraArriba Primera = new(TipoSeccionEspejoR1.Primera, 0);
    private static readonly SeccionFronteraArriba Oportuno1 = new(TipoSeccionEspejoR1.Oportuno, 1);
    private static readonly SeccionFronteraArriba Aplicacion1 = new(TipoSeccionEspejoR1.Aplicacion, 1);
    private static readonly SeccionFronteraArriba Oportuno2 = new(TipoSeccionEspejoR1.Oportuno, 2);

    [Fact]
    public void Resolver_UltimaOportuna_TomaLaUltimaSeccionPresente()
    {
        var filas = new List<FilaSeccionadaR1>
        {
            new(7, Oportuno1),
            new(13, Aplicacion1),
            new(28, Oportuno2),
        };

        var resultado = ResolvedorOperandosEmpresaR1.Resolver(filas, SeccionEsperadaR1.UltimaOportuna);

        Assert.False(resultado.EsAusente);
        Assert.False(resultado.EsAmbigua);
        Assert.Equal(28, resultado.Fila);
        Assert.Equal(Oportuno2, resultado.Seccion);
    }

    [Fact]
    public void Resolver_UnicaRecaudo_ResuelveComoLaUltimaSeccion()
    {
        // ASE5-jul-2 ENEL: UNA fila en la última sección (la única de recaudo del bloque de Totales).
        var una = new List<FilaSeccionadaR1> { new(18, Oportuno1) };
        var resultado = ResolvedorOperandosEmpresaR1.Resolver(una, SeccionEsperadaR1.UnicaRecaudo);
        Assert.Equal(18, resultado.Fila);

        // ASE5-ago-2 OCCIDENTE: 3 secciones → sigue siendo la última (nunca un fallback posicional).
        var tres = new List<FilaSeccionadaR1> { new(7, Oportuno1), new(19, Aplicacion1), new(38, Oportuno2) };
        var resultado3 = ResolvedorOperandosEmpresaR1.Resolver(tres, SeccionEsperadaR1.UnicaRecaudo);
        Assert.Equal(38, resultado3.Fila);
    }

    [Fact]
    public void Resolver_Primera_ExigeAlMenosDosSecciones()
    {
        // Con 2+ secciones: la primera (el "otro" bloque).
        var dos = new List<FilaSeccionadaR1> { new(8, Aplicacion1), new(23, Oportuno1) };
        Assert.Equal(8, ResolvedorOperandosEmpresaR1.Resolver(dos, SeccionEsperadaR1.Primera).Fila);

        // Con UNA sola sección no hay primera distinta → ausente (0 explícito, regla 6).
        var una = new List<FilaSeccionadaR1> { new(23, Oportuno2) };
        var resultado = ResolvedorOperandosEmpresaR1.Resolver(una, SeccionEsperadaR1.Primera);
        Assert.True(resultado.EsAusente);
        Assert.Null(resultado.Fila);
    }

    [Fact]
    public void Resolver_Primera_ConSeccionPropiaAusente_EsAusente()
    {
        // Oct-1 ASE3 OCCIDENTE: la empresa solo trae la sección oportuna (Oportuno#2) aunque la hoja
        // tenga otra sección con otra empresa → `Primera` (Oportuno#1) no está en su bloque → 0.
        var filas = new List<FilaSeccionadaR1> { new(23, Oportuno2) };

        Assert.True(ResolvedorOperandosEmpresaR1.Resolver(filas, SeccionEsperadaR1.Primera).EsAusente);
        Assert.Equal(23, ResolvedorOperandosEmpresaR1.Resolver(filas, SeccionEsperadaR1.UltimaOportuna).Fila);
    }

    [Fact]
    public void Resolver_Intermedia_ExigeAlMenosTresSecciones()
    {
        var dos = new List<FilaSeccionadaR1> { new(10, Oportuno1), new(26, Oportuno2) };
        Assert.True(ResolvedorOperandosEmpresaR1.Resolver(dos, SeccionEsperadaR1.Intermedia).EsAusente);

        var tres = new List<FilaSeccionadaR1> { new(10, Oportuno1), new(16, Aplicacion1), new(31, Oportuno2) };
        var resultado = ResolvedorOperandosEmpresaR1.Resolver(tres, SeccionEsperadaR1.Intermedia);
        Assert.Equal(16, resultado.Fila);
        Assert.Equal(Aplicacion1, resultado.Seccion);
    }

    [Fact]
    public void Resolver_EmpateEnLaMismaSeccion_EsSenalDeStop()
    {
        var filas = new List<FilaSeccionadaR1>
        {
            new(7, Oportuno1),
            new(12, Oportuno1),
        };

        var resultado = ResolvedorOperandosEmpresaR1.Resolver(filas, SeccionEsperadaR1.UltimaOportuna);

        Assert.True(resultado.EsAmbigua);
        Assert.Equal(Oportuno1, resultado.Seccion);
        Assert.Equal(new[] { 7, 12 }, resultado.FilasAmbiguas);
        Assert.Null(resultado.Fila);
    }

    [Fact]
    public void Resolver_BloqueVacio_EsAusente()
    {
        var resultado = ResolvedorOperandosEmpresaR1.Resolver([], SeccionEsperadaR1.UltimaOportuna);
        Assert.True(resultado.EsAusente);
    }

    [Fact]
    public void Resolver_NoUsaMagnitud_SoloSeccion()
    {
        // Magnitud ≠ ordinal (ago-2 ASE2 ENEL): el "menor" por |F| es f17 (131.543,19) pero la PRIMERA
        // sección es f7 (−1.414.859,14) → el rol `Primera` devuelve f7.
        var filas = new List<FilaSeccionadaR1>
        {
            new(7, Oportuno1),
            new(17, Aplicacion1),
            new(25, Oportuno2),
        };

        Assert.Equal(7, ResolvedorOperandosEmpresaR1.Resolver(filas, SeccionEsperadaR1.Primera).Fila);
        Assert.Equal(17, ResolvedorOperandosEmpresaR1.Resolver(filas, SeccionEsperadaR1.Intermedia).Fila);
        Assert.Equal(25, ResolvedorOperandosEmpresaR1.Resolver(filas, SeccionEsperadaR1.UltimaOportuna).Fila);
    }

    [Fact]
    public void Resolver_PrimeraSeccionPropia_NoConfundePrimeraDelBloqueConPrimeraDeLaHoja()
    {
        // Julio-2 ASE5 ENEL: una sola fila, en Oportuno#1 (la 1.ª sección de la HOJA es Aplicacion#1).
        // El bloque de la EMPRESA tiene una sola sección → `Primera` ausente (0) y `UnicaRecaudo` la toma.
        var filas = new List<FilaSeccionadaR1> { new(18, Oportuno1) };

        Assert.True(ResolvedorOperandosEmpresaR1.Resolver(filas, SeccionEsperadaR1.Primera).EsAusente);
        Assert.Equal(18, ResolvedorOperandosEmpresaR1.Resolver(filas, SeccionEsperadaR1.UnicaRecaudo).Fila);
    }

    [Fact]
    public void Resolver_SeccionInicial_PrimeraNoSeConfundeConPrimeraDeLaHoja()
    {
        // Un bloque cuya primera sección propia ES la sección sin frontera arriba (Primera#0).
        var filas = new List<FilaSeccionadaR1> { new(8, Primera), new(23, Oportuno1) };
        Assert.Equal(8, ResolvedorOperandosEmpresaR1.Resolver(filas, SeccionEsperadaR1.Primera).Fila);
    }
}
