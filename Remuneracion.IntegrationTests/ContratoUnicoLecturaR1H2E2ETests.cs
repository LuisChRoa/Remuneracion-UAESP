using System.Security.Cryptography;
using Remuneracion.Cli;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 54 (H2.2): regresión end-to-end de la ruta ÚNICA.
/// - CA1: la corrida <c>2026101</c> lee los 5 ASE SIN la excepción del defecto (el rojo era ASE3:
///   <c>"no se encontró la fila 'Aplicacion nuevos x reversion' … F223"</c>).
/// - CA4: <c>2026071</c> (Q1) y <c>2026072</c> (Q2) se resuelven POR LA MISMA RUTA y el oráculo v2 de
///   julio-2 queda intacto.
/// - CA6: las plantillas oficiales v2/v1 quedan byte-idénticas.
///
/// <para>
/// <b>Alcance del aserto (declarado, no silencioso):</b> el <c>CodigosSalida.Ok</c> de la corrida 5-ASE
/// NO se aserta porque los tres períodos terminan hoy en deudas PRE-EXISTENTES ajenas a este plan —
/// rojas ya en la línea base <c>p53-h2-suite.trx</c> (los 3 <c>ParidadCliTests</c>):
/// (a) <c>ERR-SIN-SLOT</c> en la corrida Q1: ASE2 rótulo <c>NUEVO ESQUEMA</c> sin slot compatible
/// (contrato de sección, Planes 41/46);
/// (b) <c>ERR-PLANTILLA</c> ASE5 <c>F682</c>: <see cref="AnclasExtempR1"/> lo congela como literal
/// (Plan 49 A1-interior) mientras los BYTES de la v2 lo traen como <c>&lt;f&gt;F660+F629-L629</c>. El
/// gate de anclas es **agnóstico de quincena** (no es un defecto "de Q2": muerde igual la corrida Q1;
/// p. ej. <c>Paridad_1Ase_Q1_Ase3</c>).
/// El aserto se ancla por tanto en la ETAPA DE LECTURA (los 5 ASE + coherencia multi-ASE), que es lo
/// que este plan cambia, y en los SHAs de las plantillas/oráculo.
/// </para>
/// </summary>
[Collection("Cli")]
public sealed class ContratoUnicoLecturaR1H2E2ETests
{
    private const string ShaPlantillaV2 = "0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB";
    private const string ShaPlantillaV1 = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";

    /// <summary>Mensaje del defecto que este plan elimina (rama Q1 con ordinales congelados).</summary>
    private const string MensajeDelDefecto = "no se encontró la fila 'Aplicacion nuevos x reversion'";

    private static readonly object CandadoBases = new();
    private static readonly Dictionary<string, string> Bases = new();

    private static string PlantillaV2 =>
        Ruta("Docs", "Plantilla estandarizada", "v2", "Formato Remuneracion 2026xx-x Total vf.xlsx");

    private static string PlantillaV1 =>
        Ruta("Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string OraculoJulioQ2 => Ruta(
        "Docs", "Prueba Julio-2", "Resultado", "Resultado Administrativo con plantilla estandar v2",
        "Formato Remuneracion 202607-2 Total.xlsx");

    // ════════════════════════════════════════════════════════════════════════════════════════
    // CA1 — 2026101: los 5 ASE se leen sin excepción (lector leaf, determinista)
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void S3_CA1_Octubre1_LeeLosCincoAseSinExcepcion()
    {
        var periodo = Periodo.Parse("2026101");
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var (r1, r2, r4) = Rutas("Octubre-1", aseId);
            var leaf = reader.LeerLeafInputs(AseFactory.DesdeId(aseId), periodo, r1, r2, r4);

            Assert.Equal(aseId, leaf.Ase.Id);
            Assert.NotEmpty(leaf.R1.CeldasPorAse);
            Assert.NotEmpty(leaf.R2.CeldasPorAse);
            Assert.NotEmpty(leaf.R4.CeldasPorAse);
        }

        // ASE3 (el que abortaba): ruta única, EXTEMP 0 y TOT_OPT derivado (CA3).
        var (r1Ase3, r2Ase3, r4Ase3) = Rutas("Octubre-1", 3);
        var ase3 = reader.LeerLeafInputs(AseFactory.DesdeId(3), periodo, r1Ase3, r2Ase3, r4Ase3);
        Assert.Equal(0m, ase3.R1.ExtemporaneoEsperadoPorAse);
        Assert.Equal(8811472160.34m, ase3.R1.TotalOportunoEsperadoPorAse);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // CA4 — los tres períodos pasan la ETAPA DE LECTURA por la misma ruta
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Theory]
    [InlineData("2026101", "Octubre-1")]
    [InlineData("2026071", "Julio-1")]
    [InlineData("2026072", "Julio-2")]
    public void S3_CA4_CorridaAvanzaLaLecturaDeLosCincoAse(string codigo, string carpetaPeriodo)
    {
        var shaOraculoAntes = Sha256(OraculoJulioQ2);
        var shaV2Antes = Sha256(PlantillaV2);
        var salida = Path.Combine(Path.GetTempPath(), $"p54-e2e-{codigo}-{Environment.ProcessId}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(salida);

        var stdout = new StringWriter();
        var stderr = new StringWriter();
        EjecutorCli.Ejecutar(
            ["--periodo", codigo, "--carpeta", Carpeta(carpetaPeriodo), "--plantilla", BaseDelPeriodo(codigo),
                "--salida", salida, "--cinco-ase", "--sobrescribir"],
            stdout,
            stderr);

        var log = stdout.ToString();

        // La etapa de LECTURA (la que cambia este plan) completó: los 5 ASE se leyeron y la coherencia
        // multi-ASE corrió. El código de salida NO se aserta: las deudas de escritura son externas y
        // pre-existentes (ver el doc-comment del tipo).
        Assert.Contains("ASE 5: leyendo", log, StringComparison.Ordinal);
        Assert.Contains("Validando coherencia multi-ASE", log, StringComparison.Ordinal);

        // CA1/CA7: el defecto que el plan elimina NO reaparece en ningún período.
        Assert.DoesNotContain(MensajeDelDefecto, log, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(MensajeDelDefecto, stderr.ToString(), StringComparison.OrdinalIgnoreCase);

        // CA6: ni la plantilla oficial ni el oráculo v2 de julio-2 se tocan.
        Assert.Equal(shaV2Antes, Sha256(PlantillaV2));
        Assert.Equal(shaOraculoAntes, Sha256(OraculoJulioQ2));
        Assert.Equal(ShaPlantillaV1, Sha256(PlantillaV1));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // CA6 — plantillas intactas
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void S3_CA6_PlantillasOficiales_Intactas()
    {
        Assert.Equal(ShaPlantillaV2, Sha256(PlantillaV2));
        Assert.Equal(ShaPlantillaV1, Sha256(PlantillaV1));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Base del período = plantilla oficial v2 → Phase 0 (<see cref="PreparadorBasePipeline"/>, 11
    /// tokens de período) → copia de trabajo temporal. Mismo flujo que producción.
    /// </summary>
    private static string BaseDelPeriodo(string codigo)
    {
        lock (CandadoBases)
        {
            if (Bases.TryGetValue(codigo, out var existente))
            {
                return existente;
            }

            var dir = Path.Combine(Path.GetTempPath(), $"p54-bases-{Environment.ProcessId}");
            Directory.CreateDirectory(dir);
            var salida = Path.Combine(dir, $"Plantilla_Remuneracion_{codigo}.xlsx");
            PreparadorBasePipeline.Ejecutar(PlantillaV2, Periodo.Parse(codigo), salida);
            Bases[codigo] = salida;
            return salida;
        }
    }

    private static string Carpeta(string carpetaPeriodo) => Ruta("Docs", $"Prueba {carpetaPeriodo}", "Insumos");

    private static (string R1, string R2, string R4) Rutas(string carpetaPeriodo, int aseId)
    {
        var carpeta = Path.Combine(Carpeta(carpetaPeriodo), CarpetasAse(aseId));

        return (
            Unico(carpeta, "Recaudoporcomponente"),
            Unico(carpeta, "RerpoteDetalleSaldosaFavor"),
            Unico(carpeta, "ReversiónPorComponente"));
    }

    private static string CarpetasAse(int aseId) => aseId switch
    {
        1 => "1-Promoambiental",
        2 => "2-Lime",
        3 => "3-Ciudad Limpia",
        4 => "4-Bogota Limpia",
        5 => "5-Área Limpia",
        _ => throw new ArgumentOutOfRangeException(nameof(aseId))
    };

    private static string Unico(string carpeta, string prefijo)
    {
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .SingleOrDefault(f => Path.GetFileName(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta '{prefijo}*' en {carpeta}.");
        return archivo!;
    }

    private static string Sha256(string ruta)
    {
        using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    private static string Ruta(params string[] partes) => Path.Combine(partes.Prepend(Raiz()).ToArray());
}
