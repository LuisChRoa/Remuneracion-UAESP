using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 51 (H2, S3.1): juicio A3-BANCO del <see cref="ValidadorBasico"/> — mismo grano de
/// etiquetas que el control R1/R2/R4 del Plan 50, pero sobre el reporte por banco (canal durable
/// <c>ASE n · BANCO:{label}</c>) y SIN umbral numérico.
///
/// Reglas:
/// <list type="bullet">
/// <item>etiqueta desconocida + fuente ausente en el MISMO ASE → <b>fail-fast</b> (posible fuente
/// presente no reconocida: cambio de etiqueta), con ASE + BANCO + etiquetas + fuentes;</item>
/// <item>ausencia coherente (sólo huella) → <b>verde</b> con huella (la celda queda en el 0 de la
/// plantilla);</item>
/// <item>etiqueta desconocida sin ausencia → <b>sólo evidencia</b> (no convierte en rojo lo sano).</item>
/// </list>
/// El caso base es el de <see cref="ReporteBancoTests.CrearCasoValido"/> (5 ASE reales de julio-1);
/// sólo se inyecta la huella/evidencia del bloque BANCO del ASE 2 — el resto de gates no cambia.
/// </summary>
public sealed class ControlBancoNoReconocidoPlan51Tests
{
    [Fact]
    public void A3Banco_DesconocidaConAusenciaEnElMismoAse_FailFast()
    {
        var (resultado, leafs) = ReporteBancoTests.CrearCasoValido();
        InyectarBanco(leafs[1], ["BANCO:NUEVO ESQUEMA"], ["BANCO:ESQUEMA NUEVO"]);

        var errores = new ValidadorBasico().Validar(resultado, leafs);

        var error = Assert.Single(errores, e => e.Contains("BANCO", StringComparison.Ordinal));
        Assert.Contains("ASE 2", error, StringComparison.Ordinal);
        Assert.Contains("ESQUEMA NUEVO", error, StringComparison.Ordinal);
        Assert.Contains("NUEVO ESQUEMA", error, StringComparison.Ordinal);
        Assert.Contains("Plan 51 A3", error, StringComparison.Ordinal);
    }

    [Fact]
    public void A3Banco_AusenciaCoherenteSinDesconocidas_VerdeConHuella()
    {
        var (resultado, leafs) = ReporteBancoTests.CrearCasoValido();
        InyectarBanco(leafs[1], ["BANCO:NUEVO ESQUEMA"], []);

        var errores = new ValidadorBasico().Validar(resultado, leafs);

        Assert.DoesNotContain(errores, e => e.Contains("BANCO", StringComparison.Ordinal));
        Assert.Equal(["BANCO:NUEVO ESQUEMA"], leafs[1].ReporteBanco!.Ases[0].FuentesAusentesBanco.ToArray());
    }

    [Fact]
    public void A3Banco_DesconocidaSinAusencia_EsSoloEvidencia()
    {
        var (resultado, leafs) = ReporteBancoTests.CrearCasoValido();
        InyectarBanco(leafs[1], [], ["BANCO:EAB"]);

        var errores = new ValidadorBasico().Validar(resultado, leafs);

        Assert.DoesNotContain(errores, e => e.Contains("BANCO", StringComparison.Ordinal));
        Assert.Equal(["BANCO:EAB"], leafs[1].ReporteBanco!.Ases[0].EtiquetasBancoNoReconocidas.ToArray());
    }

    [Fact]
    public void A3Banco_NoContaminaElControlR1R2R4DelPlan50()
    {
        // El control del Plan 50 lee SOLO leaf.Conciliacion: la huella BANCO no se particiona como
        // reporte R1/R2/R4 ni dispara su juicio.
        var (resultado, leafs) = ReporteBancoTests.CrearCasoValido();
        InyectarBanco(leafs[1], ["BANCO:NUEVO ESQUEMA"], ["BANCO:ESQUEMA NUEVO"]);

        var errores = new ValidadorBasico().Validar(resultado, leafs);

        Assert.DoesNotContain(errores, e => e.Contains("R1:ESQUEMA NUEVO", StringComparison.Ordinal));
        Assert.DoesNotContain(errores, e => e.Contains("· R1 · control de fuentes", StringComparison.Ordinal));
        Assert.DoesNotContain(errores, e => e.Contains("· R2 · control de fuentes", StringComparison.Ordinal));
        Assert.DoesNotContain(errores, e => e.Contains("· R4 · control de fuentes", StringComparison.Ordinal));
    }

    /// <summary>
    /// Reemplaza el bloque BANCO del leaf indicado conservando sus empresas (los gates D5 y C59
    /// siguen válidos) e inyectando la huella/evidencia que el reader produciría.
    /// </summary>
    private static void InyectarBanco(WorkbookLeafInputs leaf, string[] ausentes, string[] noReconocidas)
    {
        var original = leaf.ReporteBanco
            ?? throw new InvalidOperationException("El caso base debe traer ReporteBanco.");
        var bloque = original.Ases.Single();
        leaf.ReporteBanco = new ReporteBancoInputs
        {
            Ases =
            [
                new ReporteBancoAseInputs
                {
                    Ase = bloque.Ase,
                    Empresas = bloque.Empresas,
                    FuentesAusentesBanco = ausentes,
                    EtiquetasBancoNoReconocidas = noReconocidas
                }
            ],
            Quincena = original.Quincena,
            Consolidado = original.Consolidado
        };
    }
}
