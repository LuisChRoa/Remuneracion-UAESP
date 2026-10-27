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
    // Plan 48 (H1): la base oficial es la v2; la v1 queda historica con su propio SHA.
    private const string Sha256Oficial = "FF5CF7A819110A1E1E2CA7BDAD553DD7A25141ADD472B1904BD496F082642632";
    private const string Sha256V1Historica = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";
    private const decimal DeltaOccidenteAse5 = 417519.62m;

    [Fact]
    public void T0_Oficial_Sha256Intacto()
    {
        // Plan 48 (H1b): AMBAS plantillas intocables. v2 = base oficial vigente; v1 = histórica.
        var oficialV2 = Insumos.PlantillaOficial();
        Assert.True(File.Exists(oficialV2), $"Falta la plantilla oficial v2: {oficialV2}");
        // Plan 59 §0.6 (G-HUELLA): gate por contenido; el SHA del archivo es informativo.
        Assert.Equal(HuellaSemanticaPlantilla.V2, HuellaSemanticaPlantilla.Calcular(oficialV2));

        var oficialV1 = Path.Combine(
            Insumos.Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");
        Assert.True(File.Exists(oficialV1), $"Falta la plantilla histórica v1: {oficialV1}");
        Assert.Equal(Sha256V1Historica, Sha256(oficialV1));
    }

    [Fact]
    public void T0_Catalogo_CitaLaEvidenciaYElDeltaObservado()
    {
        // Plan 54 (H1.5): el catálogo admite filas nuevas (abierto por diseño); este T0 congela los
        // invariantes de OCCIDENTE-ASE5, que se selecciona por Id.
        var entrada = DivergenciasDeclaradasCatalogo.Entradas.Single(e => e.Id == "OCCIDENTE-ASE5");

        // La entrada CITA la evidencia T0 (R-CAT-1: entrada cita la evidencia o el test falla).
        Assert.Contains("T0-Plan41-MiniT0-OccidenteAse5", entrada.CitaEvidencia, StringComparison.Ordinal);
        Assert.Contains("T0FrenteEspejoR1Ase5EnelTests", entrada.CitaEvidencia, StringComparison.Ordinal);
        Assert.Equal(DeltaOccidenteAse5, entrada.Delta);

        // Plan 46 (H2, decisión del Ingeniero 2026-10-17): la divergencia se resolvió (Δ = 0 en la
        // corrida 9) → la entrada queda RETIRADA con fecha + motivo + evidencia citados. El
        // mecanismo (matcher/reporte/huella) queda intacto; solo cambió el DATO.
        Assert.True(entrada.Retirada);
        Assert.True(entrada.RetiroDocumentado);
        Assert.Equal("2026-10-17", entrada.FechaRetiro);
        Assert.False(entrada.Vigente);
    }

    [Fact]
    public void T0_CorridaVivaQ2_OccidenteAse5_DeltaExactoYClasificacionDeclarada()
    {
        // Plan 46 (H2, decisión del Ingeniero 2026-10-17): la divergencia real OCCIDENTE-ASE5 quedó
        // RESUELTA (F612→0; O25/D34/F34 TRUE), así que la entrada se retiró del catálogo. La corrida
        // viva end-to-end se conserva con el FIXTURE SINTÉTICO (SINTETICA-C15, Δ = 1) para congelar
        // bloqueo→reconocimiento→cierre con la MISMA maquinaria real (clasificación, reporte, naming).
        var oficial = Insumos.PlantillaOficial();
        var periodo = Insumos.PeriodoQ2();
        var dir = Path.Combine(Path.GetTempPath(), "p42-t0-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var salida = Path.Combine(dir, periodo.NombreArchivo);
        try
        {
            var procesador = FixtureDivergenciaSintetica.CrearProcesador();

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
            Assert.Equal([FixtureDivergenciaSintetica.EntradaId], clasificacion.EntradasDeclaradas);

            // 5 ASE × 1 celda (C15) = 5 veredictos declarados, todos con el Δ exacto del fixture.
            Assert.Equal(5, clasificacion.Declaradas.Count);
            foreach (var veredicto in clasificacion.Declaradas)
            {
                Assert.Equal(FixtureDivergenciaSintetica.Celda, veredicto.Celda);
                Assert.Equal(FixtureDivergenciaSintetica.Delta, veredicto.Delta);
            }

            // Plan 43 (O-4): la corrida NO cerró bien (declarada pendiente de reconocimiento) ⇒ el
            // archivo NO lleva el nombre definitivo: queda con la marca determinista "sin validar".
            // El writer SÍ produjo la salida (el gate corre después), y el oficial no se modificó.
            Assert.False(File.Exists(salida), "Una corrida no cerrada no produce el nombre definitivo.");
            Assert.True(
                File.Exists(NombresSalida.RutaSinValidar(salida)),
                "El writer debió producir la salida marcada 'sin validar' antes del gate.");
            Assert.Equal(HuellaSemanticaPlantilla.V2, HuellaSemanticaPlantilla.Calcular(oficial));
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
