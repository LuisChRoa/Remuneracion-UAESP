using Remuneracion.Core.Models;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 35 (T2, R-B-1 / D-A — regla E2): unitarios PUROS del clasificador de frontera de sección
/// <see cref="R1FirmaInterior.EsAplicacionPorFrontera"/> (sin I/O, sin OpenXML, sin períodos).
///
/// Reproduce la secuencia real del bloque ASE4-agosto (firmas A–E de <c>Recaudoporcomponente</c>,
/// verificadas por lectura directa del disco <c>Docs/Prueba Agosto-2/Insumos/4-Bogota Limpia</c>): los
/// índices 0..26 abajo son EXACTAMENTE los del bloque real. El defecto F463/H2 vivía en el predicado de
/// adyacencia (<c>Filas[i+1].EsAplicacionTotal</c>): una empresa-dato que NO cerraba su sección quedaba
/// clasificada como Oportuno aunque la próxima frontera por debajo fuese Aplicacion.
/// </summary>
public sealed class ClasificadorFronteraR1Tests
{
    private static FilaEspejoR1 F(string a = "", string b = "", string c = "", string d = "", string e = "") =>
        new() { A = a, B = b, C = c, D = d, E = e };

    /// <summary>
    /// Secuencia real ASE4-agosto (índices 0..22; los saltos intermedios son filas Vlr Servicio/Intereses
    /// y códigos D, irrelevantes para la frontera). Fronteras: MES@6, APLIC@22.
    /// </summary>
    private static IReadOnlyList<FilaEspejoR1> SecuenciaAse4Agosto() =>
    [
        F(e: "Vlr Servicio"),                 // 0
        F(d: "E", e: "Total"),                // 1
        F(c: "ENEL", d: "Total"),             // 2  dato empresa (sección Oportuno: frontera MES@6)
        F(e: "Vlr Servicio"),                 // 3
        F(d: "O", e: "Total"),                // 4
        F(c: "OCCIDENTE", d: "Total"),        // 5  dato empresa (Oportuno: frontera MES@6)
        F(b: "Mes", c: "Total"),              // 6  FRONTERA Mes
        F(a: "AFaseo", b: "Total"),           // 7
        F(e: "Vlr Servicio"),                 // 8
        F(e: "Vlr Intereses"),                // 9
        F(d: "E", e: "Total"),                // 10
        F(c: "ENEL", d: "Total"),             // 11 dato empresa (Aplicación: frontera APLIC@22)
        F(e: "Vlr Servicio"),                 // 12
        F(e: "Vlr Intereses"),                // 13
        F(d: "1", e: "Total"),                // 14
        F(e: "Vlr Servicio"),                 // 15
        F(d: "5", e: "Total"),                // 16
        F(c: "NUEVO ESQUEMA", d: "Total"),    // 17 dato empresa (Aplicación: frontera APLIC@22)
        F(e: "Vlr Servicio"),                 // 18
        F(e: "Vlr Intereses"),                // 19
        F(d: "O", e: "Total"),                // 20
        F(c: "OCCIDENTE", d: "Total"),        // 21 dato empresa (Aplicación: frontera inmediata APLIC@22)
        F(b: "Aplicacion nuevos x reversion", c: "Total") // 22 FRONTERA Aplicación
    ];

    [Theory]
    [InlineData(2, false)]   // ENEL en sección Oportuno (frontera MES@6)
    [InlineData(5, false)]   // OCCIDENTE en sección Oportuno (frontera MES@6)
    [InlineData(11, true)]   // ENEL no cierra su sección: la próxima frontera es APLIC@22 (H2)
    [InlineData(17, true)]   // NUEVO ESQUEMA: próxima frontera APLIC@22 (H2)
    [InlineData(21, true)]   // OCCIDENTE con frontera inmediata APLIC@22 (caso base intacto)
    public void Ase4Agosto_FronteraClasificaComoElManual(int indice, bool esAplicacion)
    {
        Assert.Equal(esAplicacion, R1FirmaInterior.EsAplicacionPorFrontera(SecuenciaAse4Agosto(), indice));
    }

    [Fact]
    public void FronteraEsLaProximaMesOAplic_NoElVecinoInmediato()
    {
        // dato (0), ruido intermedio, APLIC (2), Mes más abajo (5): gana la próxima (APLIC@2).
        var filas = new List<FilaEspejoR1>
        {
            F(c: "ENEL", d: "Total"),                 // 0 dato
            F(e: "Vlr Servicio"),                     // 1
            F(b: "Aplicacion nuevos x reversion", c: "Total"), // 2 APLIC
            F(e: "Vlr Servicio"),                     // 3
            F(e: "Vlr Intereses"),                    // 4
            F(b: "Mes", c: "Total")                   // 5 MES
        };

        Assert.True(R1FirmaInterior.EsAplicacionPorFrontera(filas, 0));
    }

    [Fact]
    public void FronteraMesMasCercana_ClasificaOportunoAunqueHayaAplicMasAbajo()
    {
        var filas = new List<FilaEspejoR1>
        {
            F(c: "ENEL", d: "Total"),                 // 0 dato
            F(e: "Vlr Servicio"),                     // 1
            F(b: "Mes", c: "Total"),                  // 2 MES (más cercana)
            F(b: "Aplicacion nuevos x reversion", c: "Total") // 3 APLIC
        };

        Assert.False(R1FirmaInterior.EsAplicacionPorFrontera(filas, 0));
    }

    [Fact]
    public void SinFronteraPorDebajo_DevuelveNull_NuncaCeroSilencioso()
    {
        var filas = new List<FilaEspejoR1> { F(c: "ENEL", d: "Total"), F(e: "Vlr Servicio") };

        Assert.Null(R1FirmaInterior.EsAplicacionPorFrontera(filas, 0));
    }

    [Fact]
    public void Julio_Identidad_CasoBaseAdyacente_ConservaComportamiento()
    {
        // Geometría de julio (sin recorte): la empresa-dato de Aplicación cierra su sección de forma
        // adyacente. La regla nueva debe reproducir la clasificación histórica.
        var filas = new List<FilaEspejoR1>
        {
            F(c: "ENEL", d: "Total"),                 // 0 dato Aplicación (adyacente)
            F(b: "Aplicacion nuevos x reversion", c: "Total"), // 1 APLIC
            F(c: "OCCIDENTE", d: "Total"),            // 2 dato Oportuno
            F(b: "Mes", c: "Total")                   // 3 MES
        };

        Assert.True(R1FirmaInterior.EsAplicacionPorFrontera(filas, 0));
        Assert.False(R1FirmaInterior.EsAplicacionPorFrontera(filas, 2));
    }
}
