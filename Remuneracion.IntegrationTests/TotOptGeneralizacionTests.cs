using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 25 (WU-2 = T3, punto A / R-F-5): guardián de la GENERALIZACIÓN del <c>totOpt</c> por
/// firma sobre la secuencia espejo.
///
/// El mapa congelado Q2 de ASE5 asumía 2 filas <c>Mes/Total</c> (variante de julio). La fuente de
/// agosto trae 3 (T0a §2.1) y la fuente de julio trae 2: la fórmula visible del template —
/// <c>ΣF(todas las Mes) − Especiales de todas MENOS la última</c> — debe dar EXACTAMENTE el mismo
/// resultado que el mapa congelado en julio (guardián bit-comparable) y el valor generalizado en
/// agosto, sin depender del conteo.
///
/// Valores esperados (oráculo de julio: caché golden D13; agosto: fórmula visible T0b/T0c con las
/// 3 filas Mes de la fuente real). Sin fixtures sintéticas.
/// </summary>
public sealed class TotOptGeneralizacionTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    /// <summary>
    /// ASE5-julio: 2 filas <c>Mes/Total</c>. Guardián de NO-regresión: el <c>totOpt</c> por firma
    /// debe ser idéntico al mapa congelado que cerró el golden (12.033.011.685,71 = D13).
    /// </summary>
    [Fact]
    public void TotOpt_Ase5Julio_2FilasMes_IdenticoAlMapaCongelado()
    {
        var leaf = new ExcelDataReaderWorkbookLeafInputReader().LeerLeafInputs(
            Insumos.Ase(5), Insumos.PeriodoQ2(), Insumos.R1Q2(5), Insumos.R2Q2(5), Insumos.R4Q2(5));

        Assert.InRange(leaf.R1.TotalOportunoEsperadoPorAse - 12033011685.71m, -Tolerancia, Tolerancia);
    }

    /// <summary>
    /// ASE5-agosto: 3 filas <c>Mes/Total</c> (fuente real). <c>totOpt</c> generalizado = 12.105.458.586,04
    /// (misma fórmula visible que julio, con la ocurrencia extra). Cierra el DetRetri de ASE5
    /// contra el R10 (12.137.660.178) en la regresión end-to-end.
    /// </summary>
    [Fact]
    public void TotOpt_Ase5Agosto_3FilasMes_GeneralizadoCierra()
    {
        var leaf = new ExcelDataReaderWorkbookLeafInputReader().LeerLeafInputs(
            Insumos.Ase(5), Insumos.PeriodoQ2(), Insumos.R1Agosto(5), Insumos.R2Agosto(5), Insumos.R4Agosto(5));

        Assert.InRange(leaf.R1.TotalOportunoEsperadoPorAse - 12105458586.04m, -Tolerancia, Tolerancia);
    }

    /// <summary>
    /// Punto A explícito: el MISMO ASE (5), la MISMA firma, pero 2 filas (julio) vs 3 filas
    /// (agosto) → cada uno contra su valor esperado propio. La resolución por firma absorbe la
    /// cardinalidad sin mapa nuevo (OCP del Plan 25).
    /// </summary>
    [Fact]
    public void TotOpt_Ase5_MismaFirmaDistintaCardinalidad_AmbosContraSuValor()
    {
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        var julio = reader.LeerLeafInputs(
            Insumos.Ase(5), Insumos.PeriodoQ2(), Insumos.R1Q2(5), Insumos.R2Q2(5), Insumos.R4Q2(5));
        var agosto = reader.LeerLeafInputs(
            Insumos.Ase(5), Insumos.PeriodoQ2(), Insumos.R1Agosto(5), Insumos.R2Agosto(5), Insumos.R4Agosto(5));

        Assert.InRange(julio.R1.TotalOportunoEsperadoPorAse - 12033011685.71m, -Tolerancia, Tolerancia);
        Assert.InRange(agosto.R1.TotalOportunoEsperadoPorAse - 12105458586.04m, -Tolerancia, Tolerancia);

        // La cardinalidad de filas Mes difiere (2 vs 3) → los totOpt son distintos por dato, no por lógica.
        Assert.NotEqual(julio.R1.TotalOportunoEsperadoPorAse, agosto.R1.TotalOportunoEsperadoPorAse);
    }
}
