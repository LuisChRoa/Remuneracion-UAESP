using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 46 (H2, decisión del Ingeniero 2026-10-17 — opción A): FIXTURE SINTÉTICO de divergencia
/// declarada. La divergencia real <c>OCCIDENTE-ASE5</c> ya no existe (el fix la resolvió), pero el
/// MECANISMO de divergencia declarada (bloqueo → reconocimiento → cierre + reporte + huella +
/// sellado) sigue siendo obligatorio y debe conservar su cobertura end-to-end.
///
/// Cómo se sintetiza (sin tocar el mecanismo ni la plantilla oficial):
/// <list type="bullet">
/// <item><b>Hecho real en el workbook:</b> <see cref="ReaderConRecaudoInconsistente"/> decora el
/// lector y suma +1 al <c>Total</c> del primer <see cref="RecaudoEmpresaInputs"/> SIN tocar sus
/// celdas. El writer puebla <c>'Valida - Control Recaudo'!F10 = Σ Total</c> y las hojas
/// <c>Recaudo *</c> con las celdas → <c>VALIDACION_TOTAL!C9 = SUM(C3:C8)</c> queda en Σ celdas y
/// <c>C15 = (C14 = C9)</c> evalúa FALSE con <b>Δ = 1</b> exacto; el sellado (read-back) escribe 0.</item>
/// <item><b>Entrada declarada sintética:</b> <see cref="Catalogo"/> contiene una entrada COMPLETA y
/// VIGENTE (<c>SINTETICA-C15</c>, celda <c>C15</c>, Δ = 1) con dueño/cita/condición de retiro.</item>
/// <item><b>Clasificación:</b> se ejercita el MISMO <see cref="ValidadorBasico"/> de producción,
/// instanciado con el catálogo sintético (costura de constructor; review W-1) — <b>cero lógica
/// duplicada</b> (antes el fixture espejaba la clasificación línea a línea).</item>
/// </list>
/// El catálogo REAL (<see cref="DivergenciasDeclaradasCatalogo"/>) NO se contamina: el fixture usa
/// la sobrecarga pura <c>Buscar(catalogo, celda, observado, esperado)</c>. El id sintético no está
/// en el catálogo real, así que el reporte/UI muestran su fallback administrativo documentado
/// ("responsable externo por confirmar") — el resto del copy (celda, Δ, aceptación) es real.
/// </summary>
internal static class FixtureDivergenciaSintetica
{
    /// <summary>Id de la entrada sintética (no existe en el catálogo real).</summary>
    public const string EntradaId = "SINTETICA-C15";

    /// <summary>Celda del sub-bloque FALSE sintetizado (<c>C14 = C9</c>).</summary>
    public const string Celda = "C15";

    /// <summary>Δ exacto sintetizado (+1 al total de recaudo de una empresa vs sus celdas).</summary>
    public const decimal Delta = 1m;

    /// <summary>Catálogo sintético (una sola entrada COMPLETA y VIGENTE).</summary>
    public static IReadOnlyList<DivergenciasDeclaradasCatalogo.Entrada> Catalogo { get; } =
    [
        new(
            Id: EntradaId,
            Celdas: [Celda],
            Delta: Delta,
            OrigenMonto:
                "Fixture sintético Plan 46 H2: el Total del primer 'Recaudo *' se incrementa en +1 " +
                "respecto de sus celdas; VALIDACION_TOTAL!C14 (=Σ Total) deja de igualar C9 (=SUM de " +
                "las celdas) con Δ exacto = 1.",
            Dueno: "PRUEBA (fixture sintético)",
            CitaEvidencia:
                "Remuneracion.IntegrationTests/FixtureDivergenciaSintetica.cs + " +
                "plans/46 (H2, decisión del Ingeniero 2026-10-17, opción A).",
            Fecha: "2026-10-17",
            CondicionRetiro: "Se retira cuando el fixture sintético deje de usarse.")
    ];

    /// <summary>
    /// Validador REAL (<see cref="ValidadorBasico"/>) con el catálogo sintético. Plan 46 (review W-1):
    /// ya NO se espeja la clasificación — se ejercita la MISMA implementación de producción con un
    /// catálogo distinto (costura de constructor; default de producción = catálogo real).
    /// </summary>
    public static IValidador CrearValidador() => new ValidadorBasico(Catalogo);

    /// <summary>Lector de leaf con la inconsistencia sintética (Total +1 vs celdas, Δ = 1).</summary>
    public static IWorkbookLeafInputReader CrearLector() => new ReaderConRecaudoInconsistente();

    /// <summary>Procesador de período con el hecho sintético ya cableado.</summary>
    public static ProcesadorPeriodo CrearProcesador(
        bool conOraculo = true,
        bool conReporte = true,
        IWorkbookLeafInputReader? lector = null) =>
        new(
            new ExcelDataReaderRecaudoReader(),
            lector ?? new ReaderConRecaudoInconsistente(),
            new CalculoRemuneracion(),
            CrearValidador(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader(),
            conOraculo ? new ValidacionOracleReader() : null,
            conReporte ? new ReporteDivergenciasDeclaradasWriter() : null);

    /// <summary>
    /// Lector decorador: delega TODO en el lector real salvo <c>LeerRecaudosEmpresa</c>, donde suma
    /// +1 al <c>Total</c> del primer recaudo (inconsistencia sintética, Δ = 1). No toca <c>Celdas</c>:
    /// el resto del workbook queda idéntico al de una corrida sana.
    /// </summary>
    private sealed class ReaderConRecaudoInconsistente : IWorkbookLeafInputReader
    {
        private readonly ExcelDataReaderWorkbookLeafInputReader _inner = new();

        public WorkbookLeafInputs LeerLeafInputs(Ase ase, Periodo periodo, string rutaR1, string rutaR2, string rutaR4) =>
            _inner.LeerLeafInputs(ase, periodo, rutaR1, rutaR2, rutaR4);

        public IReadOnlyList<ConciliacionEmpresaInputs> LeerConciliacionEmpresas(Ase ase, Periodo periodo, string rutaR1, string rutaR2, string rutaR4) =>
            _inner.LeerConciliacionEmpresas(ase, periodo, rutaR1, rutaR2, rutaR4);

        public IReadOnlyList<RecaudoEmpresaInputs> LeerRecaudosEmpresa(Periodo periodo, Func<EmpresaFacturacion, string?> rutaConciliacionPorEmpresa)
        {
            var recaudos = _inner.LeerRecaudosEmpresa(periodo, rutaConciliacionPorEmpresa).ToList();
            if (recaudos.Count > 0)
            {
                recaudos[0].Total += Delta;
            }

            return recaudos;
        }

        public ReporteBancoInputs LeerReporteBanco(Ase ase, Periodo periodo, string rutaReportePagosxBanco) =>
            _inner.LeerReporteBanco(ase, periodo, rutaReportePagosxBanco);

        public BalanceScInputs LeerBalanceSc(Ase ase, Periodo periodo, string rutaBalance) =>
            _inner.LeerBalanceSc(ase, periodo, rutaBalance);

        public SaldosNotasAseInputs LeerSaldosNotas(Ase ase, string rutaSaldosNotas) =>
            _inner.LeerSaldosNotas(ase, rutaSaldosNotas);

        public RetribucionNegativaAseInputs LeerRetribucionNegativa(Ase ase, string rutaRetribucionNegativa) =>
            _inner.LeerRetribucionNegativa(ase, rutaRetribucionNegativa);

        public BloqueEspejoAseInputs LeerEspejoR1(Ase ase, string rutaR1) =>
            _inner.LeerEspejoR1(ase, rutaR1);

        public DetalleR2AseInputs LeerDetalleR2(Ase ase, string rutaR2) =>
            _inner.LeerDetalleR2(ase, rutaR2);

        public DetalleR4AseInputs LeerDetalleR4(Ase ase, string rutaR4) =>
            _inner.LeerDetalleR4(ase, rutaR4);
    }
}
