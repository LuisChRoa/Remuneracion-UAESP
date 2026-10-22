using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 50 (H2) — A3, control anti-falso-0 (S3.1).
///
/// El T0 de H0 (`Docs/T0-Plan50-FuentesRecaudoOpcionales.md` §3) demostró con bytes que la vía
/// aritmética "Σ fuentes presentes vs total del reporte" NO es viable en R1 (el total del cuadro
/// Componente no es la suma de las filas-etiqueta) ni en R2 (el total incluye <c>EAB</c>, fuente real
/// sin destino en el mapa). El control mínimo alternativo (§2.5) es el de ETIQUETAS:
///
/// - etiqueta observada que no resuelve a ninguna empresa del mapa = EVIDENCIA
///   (<see cref="ConciliacionEmpresaInputs.EtiquetasNoReconocidas"/>);
/// - evidencia + fuente ausente EN EL MISMO REPORTE del mismo ASE = fail-fast (posible fuente
///   presente no reconocida: hay dinero que ninguna fuente reconocida explica);
/// - evidencia sin ausencias (o con ausencias de OTRO reporte) NO dispara.
/// </summary>
public sealed class ControlAntiFalsoCeroPlan50Tests
{
    [Fact]
    public void S3_1_EtiquetaNoReconocidaSinAusencias_EsSoloEvidencia()
    {
        var (resultado, leafs) = ConciliacionEmpresaTests.CrearCasoValido();
        foreach (var conc in leafs[1].Conciliacion)
        {
            conc.EtiquetasNoReconocidas = ["R2:EAB"];
        }

        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    [Fact]
    public void S3_1_EtiquetaNoReconocidaConAusenciaEnElMismoReporte_DisparaConDetalle()
    {
        var (resultado, leafs) = ConciliacionEmpresaTests.CrearCasoValido();
        foreach (var conc in leafs[1].Conciliacion)
        {
            conc.EtiquetasNoReconocidas = ["R2:EAB"];
        }

        leafs[1].Conciliacion.Single(c => c.Empresa.Id == 4).FuentesAusentes = ["R2:OCCIDENTE"];

        var errores = new ValidadorBasico().Validar(resultado, leafs);

        var error = Assert.Single(errores);
        Assert.Contains("[ERR-VALIDACION]", error, StringComparison.Ordinal);
        Assert.Contains("ASE 2", error, StringComparison.Ordinal);
        Assert.Contains("R2", error, StringComparison.Ordinal);
        Assert.Contains("EAB", error, StringComparison.Ordinal);
        Assert.Contains("OCCIDENTE", error, StringComparison.Ordinal);
        Assert.Contains("Plan 50 A3", error, StringComparison.Ordinal);
    }

    [Fact]
    public void S3_1_EtiquetaNoReconocidaConAusenciaEnOtroReporte_NoDispara()
    {
        // Contracara del caso real de octubre-1 ASE2: `EAB` (R2) convive con ausencias de R1/R4 del
        // MISMO ASE — el control se acota por reporte, así que no hay falso rojo.
        var (resultado, leafs) = ConciliacionEmpresaTests.CrearCasoValido();
        foreach (var conc in leafs[1].Conciliacion)
        {
            conc.EtiquetasNoReconocidas = ["R2:EAB"];
        }

        var recip = leafs[1].Conciliacion.Single(c => c.Empresa.Id == 1);
        recip.FuentesAusentes = ["R1:NUEVO ESQUEMA", "R4:NUEVO ESQUEMA"];

        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    [Fact]
    public void S3_1_AusenciaCoherenteSinEtiquetas_NoDispara()
    {
        // Ausencia total de una fuente con el control de etiquetas limpio = caso normal (0 + huella).
        var (resultado, leafs) = ConciliacionEmpresaTests.CrearCasoValido();
        leafs[1].Conciliacion.Single(c => c.Empresa.Id == 1).FuentesAusentes = ["R1:NUEVO ESQUEMA"];

        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    [Fact]
    public void S3_1_EvidenciaReal_OctubreAse2_ReportaEabSinDisparar()
    {
        // Evidencia del reader sobre la fuente real: `EAB` (R2) es la única etiqueta no reconocida y
        // convive con ausencias de OTROS reportes (R1/R4) → el control no dispara en la corrida sana.
        var conciliacion = new ExcelDataReaderWorkbookLeafInputReader().LeerConciliacionEmpresas(
            Insumos.Ase(2),
            Periodo.Parse("2026101"),
            RutaFuente("Prueba Octubre-1", 2, "Recaudoporcomponente"),
            RutaFuente("Prueba Octubre-1", 2, "RerpoteDetalleSaldosaFavor"),
            RutaFuente("Prueba Octubre-1", 2, "Reversi"));

        Assert.All(conciliacion, conc => Assert.Contains("R2:EAB", conc.EtiquetasNoReconocidas));
        Assert.All(conciliacion, conc => Assert.DoesNotContain(
            conc.EtiquetasNoReconocidas,
            etiqueta => etiqueta.StartsWith("R1:", StringComparison.OrdinalIgnoreCase) || etiqueta.StartsWith("R4:", StringComparison.OrdinalIgnoreCase)));

        var ausentes = conciliacion.SelectMany(c => c.FuentesAusentes).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        Assert.Equal(new[] { "R1:NUEVO ESQUEMA", "R4:NUEVO ESQUEMA" }, ausentes.OrderBy(x => x, StringComparer.Ordinal));
    }

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md"))) return dir.FullName;
            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    private static string RutaFuente(string periodo, int ase, string prefijo)
    {
        var dir = Path.Combine(Raiz(), "Docs", periodo, "Insumos", CarpetasAse[ase - 1]);
        Assert.True(Directory.Exists(dir), $"Falta la carpeta explícita de insumos: {dir}");
        var archivo = Directory.EnumerateFiles(dir, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta el reporte '{prefijo}*' del ASE {ase} en {dir}");
        return archivo!;
    }
}
