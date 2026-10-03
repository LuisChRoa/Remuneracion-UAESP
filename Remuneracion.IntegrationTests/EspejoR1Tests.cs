using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 21 — PR 2 (T1 + T2): modelo espejo R1 y lector espejo.
///
/// Evidencia congelada: <c>plans/21 - T0 Evidencia.md</c> §6 (T0e). Los conteos por ASE
/// (Q1 38/52/46/69/44; Q2 45/75/43/73/52; agosto 42/66/37/79/60) y las tres invariantes duras
/// (<c>Componente/Total</c>, <c>Subs/Cont/Total</c>, <c>Total</c> final) se verifican contra las
/// fuentes reales. Agosto NO tiene salida/plantilla: su verificación es estructural (smoke).
/// </summary>
public sealed class EspejoR1Tests
{
    private static readonly int[] ConteoQ1 = [38, 52, 46, 69, 44];
    private static readonly int[] ConteoQ2 = [45, 75, 43, 73, 52];
    private static readonly int[] ConteoAgosto = [42, 66, 37, 79, 60];

    [Fact]
    public void Modelo_FilaEspejoR1_FirmaYPredicadosDeCierre()
    {
        var componente = new FilaEspejoR1
        {
            A = "Componente",
            B = "Total",
            ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase)
            {
                ["Total"] = 19556118465.99m,
                ["SERVICIO ESPECIALES"] = 18193739.24m
            }
        };

        Assert.Equal("Componente|Total|||", componente.Firma);
        Assert.True(componente.EsComponenteTotal);
        Assert.False(componente.EsSubsContTotal);
        Assert.False(componente.EsTotalFinal);
        Assert.Equal(19556118465.99m, componente.Valor("Total"));
        Assert.Equal(18193739.24m, componente.Valor("SERVICIO ESPECIALES"));
        Assert.Null(componente.Valor("NoExiste"));

        // La columna "SERVICIO ESPECIALES" se busca por encabezado (case-insensitive).
        Assert.Equal(18193739.24m, componente.Valor("servicio especiales"));

        var subs = new FilaEspejoR1 { A = "Subs/Cont", B = "Total" };
        Assert.True(subs.EsSubsContTotal);
        Assert.False(subs.EsComponenteTotal);

        var total = new FilaEspejoR1 { A = "Total" };
        Assert.True(total.EsTotalFinal);
        Assert.False(total.EsComponenteTotal);
    }

    [Fact]
    public void T0e_Q1_SecuenciaEspejoPorAse_ConteosEInvariantes()
        => VerificarPeriodo(1, Insumos.R1, ConteoQ1);

    [Fact]
    public void T0e_Q2_SecuenciaEspejoPorAse_ConteosEInvariantes()
        => VerificarPeriodo(2, Insumos.R1Q2, ConteoQ2);

    [Fact]
    public void T0e_Agosto_SmokeSecuenciaEspejoPorAse_ConteosEInvariantes()
        => VerificarPeriodo(0, AgostoR1, ConteoAgosto);

    [Fact]
    public void T0e_Q1_Ase1_ComponenteTotal_ValoresPorHeader_ContraFuente()
    {
        // Evidencia T0 (dump OpenXML de la fuente ASE1-Q1, fila R25):
        //   A=Componente B=Total F=19556118465.99 L(SERVICIO ESPECIALES)=18193739.24.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var bloque = reader.LeerEspejoR1(Insumos.Ase(1), Insumos.R1(1));

        var fila = bloque.ComponenteTotal;
        Assert.NotNull(fila);
        Assert.InRange((fila!.Valor("Total") ?? 0m) - 19556118465.99m, -Insumos.Tolerancia, Insumos.Tolerancia);
        Assert.InRange((fila.Valor("SERVICIO ESPECIALES") ?? 0m) - 18193739.24m, -Insumos.Tolerancia, Insumos.Tolerancia);
        Assert.Equal("Componente", fila.A);
        Assert.Equal("Total", fila.B);
    }

    /// <summary>
    /// W-6 (auditoría PR3, R-E-5/S7): si la fuente del período no trae una de las 3 invariantes
    /// duras T0e, el fail-fast nombra ASE + reporte + fila esperada. Se parte de una fuente REAL
    /// (Q2 ASE2) y se quita la fila invariante EN MEMORIA (sin fixtures sintéticas en disco),
    /// ejercitando directamente <c>ValidarInvariantesEspejoR1</c> (W-7: logica unificada).
    /// </summary>
    [Theory]
    [InlineData("componente")]
    [InlineData("subscont")]
    [InlineData("totalfinal")]
    public void T0e_FuenteSinInvariante_FallaNombrandoAseReporteFila(string invariante)
    {
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var real = reader.LeerEspejoR1(Insumos.Ase(2), Insumos.R1Q2(2));

        var filasSinInvariante = real.Filas.Where(f => invariante switch
        {
            "componente" => !f.EsComponenteTotal,
            "subscont" => !f.EsSubsContTotal,
            _ => !f.EsTotalFinal
        }).ToList();

        var mutado = new BloqueEspejoAseInputs
        {
            Ase = real.Ase,
            Encabezados = real.Encabezados,
            TieneColumnaEspeciales = real.TieneColumnaEspeciales,
            Filas = filasSinInvariante
        };

        var etiquetaEsperada = invariante switch
        {
            "componente" => "Componente/Total",
            "subscont" => "Subs/Cont/Total",
            _ => "Total (A='Total', B vacío)"
        };

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            ProcesadorPeriodo.ValidarInvariantesEspejoR1(real.Ase, mutado));

        Assert.Contains("ASE 2", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Reporte Componentes R1", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(etiquetaEsperada, ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static void VerificarPeriodo(int quincena, Func<int, string> rutaPorAse, int[] conteoEsperado)
    {
        var etiqueta = quincena switch { 1 => "Q1-julio", 2 => "Q2-julio", _ => "agosto" };
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var bloque = reader.LeerEspejoR1(Insumos.Ase(aseId), rutaPorAse(aseId));

            Assert.Equal(aseId, bloque.Ase.Id);
            Assert.Equal(conteoEsperado[aseId - 1], bloque.TotalFilas);
            Assert.Equal(bloque.TotalFilas, bloque.Filas.Count);

            // Encabezados detectados dinámicamente: la cantidad VARÍA por ASE (p. ej. ASE2-Q1 trae
            // 16 columnas F..U con "Rec.Intereses"); solo los encabezados núcleo son comunes.
            Assert.True(bloque.Encabezados.Count >= 15, $"{etiqueta} ASE{aseId}: headers({bloque.Encabezados.Count}) = {string.Join(" | ", bloque.Encabezados)}");
            Assert.Contains("Total", bloque.Encabezados);
            Assert.Contains("SERVICIO ESPECIALES", bloque.Encabezados);
            Assert.Contains("Componente TDF", bloque.Encabezados);
            Assert.True(bloque.TieneColumnaEspeciales, $"{etiqueta} ASE{aseId}: la fuente trae SERVICIO ESPECIALES.");

            // Invariantes duras T0e.
            Assert.True(bloque.TieneInvariantesDeCierre, $"{etiqueta} ASE{aseId}: faltan invariantes de cierre T0e.");
            Assert.NotNull(bloque.ComponenteTotal);
            Assert.NotNull(bloque.SubsContTotal);
            Assert.NotNull(bloque.TotalFinal);

            // La secuencia observada es completa: toda fila aporta la columna Total (F).
            Assert.All(bloque.Filas, fila => Assert.NotNull(fila.Valor("Total")));
        }
    }

    private static string AgostoR1(int aseId)
    {
        var carpeta = Path.Combine(Insumos.Raiz(), "Docs", "Prueba2", "Insumos");
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta de agosto: {carpeta}");
        var dir = Directory.EnumerateDirectories(carpeta)
            .First(d => Path.GetFileName(d).StartsWith($"{aseId}-", StringComparison.OrdinalIgnoreCase));
        return Directory.EnumerateFiles(dir, "Recaudoporcomponente*.xlsx", SearchOption.TopDirectoryOnly).First();
    }
}
