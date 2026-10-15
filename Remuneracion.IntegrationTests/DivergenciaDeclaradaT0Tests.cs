using System.Security.Cryptography;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 42 (T0, R-T-1, D-A): evidencia congelada de la primera (y hoy única) entrada del catálogo de
/// divergencias declaradas — OCCIDENTE ASE5 (oficial 3 términos vs manual legacy 2 términos,
/// Δ = 417.519,62). Fixture read-only: corre el path end-to-end desde la plantilla OFICIAL a una
/// ruta temp FRESCA y congela (i) el Δ exacto en O25/D34/F34, (ii) la clasificación Declarada con su
/// referencia, (iii) el vínculo entrada↔evidencia, y (iv) el SHA256 del oficial intacto.
///
/// Documento que manda: <c>Docs/T0-Plan42-Taxonomia-DivergenciasDeclaradas.md</c>; la geometría de la
/// fórmula está congelada en <c>T0FrenteEspejoR1Ase5EnelTests</c> (Plan 41). NO toca producción.
/// </summary>
public sealed class DivergenciaDeclaradaT0Tests
{
    private const string Sha256Oficial = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";
    private const decimal DeltaOccidenteAse5 = 417519.62m;

    [Fact]
    public void T0_Oficial_Sha256Intacto()
    {
        var oficial = Insumos.PlantillaOficial();
        Assert.True(File.Exists(oficial), $"Falta la plantilla oficial: {oficial}");
        Assert.Equal(Sha256Oficial, Sha256(oficial));
    }

    [Fact]
    public void T0_Catalogo_CitaLaEvidenciaYElDeltaObservado()
    {
        var entrada = Assert.Single(DivergenciasDeclaradasCatalogo.Entradas);

        // La entrada CITA la evidencia T0 (R-CAT-1: entrada cita la evidencia o el test falla).
        Assert.Contains("T0-Plan41-MiniT0-OccidenteAse5", entrada.CitaEvidencia, StringComparison.Ordinal);
        Assert.Contains("T0FrenteEspejoR1Ase5EnelTests", entrada.CitaEvidencia, StringComparison.Ordinal);
        Assert.Equal(DeltaOccidenteAse5, entrada.Delta);
    }

    [Fact]
    public void T0_CorridaVivaQ2_OccidenteAse5_DeltaExactoYClasificacionDeclarada()
    {
        var oficial = Insumos.PlantillaOficial();
        var periodo = Insumos.PeriodoQ2();
        var dir = Path.Combine(Path.GetTempPath(), "p42-t0-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var salida = Path.Combine(dir, periodo.NombreArchivo);
        try
        {
            var procesador = new ProcesadorPeriodo(
                new ExcelDataReaderRecaudoReader(),
                new ExcelDataReaderWorkbookLeafInputReader(),
                new CalculoRemuneracion(),
                new ValidadorBasico(),
                new OpenXmlPlantillaWriter(),
                new ArchivoFuenteLocator(),
                new ExcelDataReaderDetRetriR10Reader(),
                new ValidacionOracleReader());

            // Sin reconocimiento explícito: la divergencia declarada BLOQUEA (S3) y porta la
            // clasificación (la UI/CLI la presentan sin re-evaluar).
            var ex = Assert.Throws<CalculoInvalidoException>(() => procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = periodo,
                CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                RutaPlantilla = oficial,
                RutaSalida = salida
            }));

            var clasificacion = ex.Clasificacion;
            Assert.NotNull(clasificacion);
            Assert.True(clasificacion!.SoloDeclaradas);
            Assert.False(clasificacion.HayNoDeclaradas);
            Assert.Equal(["OCCIDENTE-ASE5"], clasificacion.EntradasDeclaradas);

            // 5 ASE × O25/D34/F34 = 15 veredictos declarados, todos con el Δ exacto de la evidencia.
            Assert.Equal(15, clasificacion.Declaradas.Count);
            foreach (var veredicto in clasificacion.Declaradas)
            {
                Assert.Contains(veredicto.Celda, new[] { "O25", "D34", "F34" });
                Assert.Equal(DeltaOccidenteAse5, veredicto.Delta);
            }

            // Plan 43 (O-4): la corrida NO cerró bien (declarada pendiente de reconocimiento) ⇒ el
            // archivo NO lleva el nombre definitivo: queda con la marca determinista "sin validar".
            // El writer SÍ produjo la salida (el gate corre después), y el oficial no se modificó.
            Assert.False(File.Exists(salida), "Una corrida no cerrada no produce el nombre definitivo.");
            Assert.True(
                File.Exists(NombresSalida.RutaSinValidar(salida)),
                "El writer debió producir la salida marcada 'sin validar' antes del gate.");
            Assert.Equal(Sha256Oficial, Sha256(oficial));
        }
        finally
        {
            try
            {
                Directory.Delete(dir, recursive: true);
            }
            catch
            {
                // best-effort
            }
        }
    }

    [Fact]
    public void T0_Taxonomia_CatalogoSoloTieneDivergenciasExternas()
    {
        // La taxonomía T0 cierra la lista de declarables: ninguna entrada cita códigos internos
        // (lectura/cálculo/colocación/sellado/naming/preflight) — esos NUNCA son declarables.
        string[] codigosInternos = ["ERR-SIN-SLOT", "ERR-PLANTILLA", "ERR-FORMATO-FUENTE", "ERR-FUENTE-NO-ENCONTRADA"];
        foreach (var entrada in DivergenciasDeclaradasCatalogo.Entradas)
        {
            foreach (var codigo in codigosInternos)
            {
                Assert.DoesNotContain(codigo, entrada.CitaEvidencia, StringComparison.Ordinal);
                Assert.DoesNotContain(codigo, entrada.OrigenMonto, StringComparison.Ordinal);
            }

            // Toda entrada tiene un dueño EXTERNO a la aplicación (no un gate ni un componente interno).
            Assert.False(string.IsNullOrWhiteSpace(entrada.Dueno));
        }
    }

    private static string Sha256(string ruta) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(ruta)));
}
