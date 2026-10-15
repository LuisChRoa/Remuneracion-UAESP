using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;

namespace Remuneracion.WinForms;

/// <summary>
/// Plan 43 (O-1): composición ÚNICA de los servicios del host WinForms. Es PURA (solo Core +
/// Infrastructure, sin tipos de UI), por lo que los tests la compilan por archivo vinculado y
/// ejercitan la MISMA unidad que arranca la aplicación.
///
/// El oráculo de validaciones cruzadas (<see cref="ValidacionOracleReader"/>) y el emisor del
/// reporte durable (<see cref="ReporteDivergenciasDeclaradasWriter"/>) son SIEMPRE parte de la
/// composición: ninguna vía de arranque —incluido el constructor por defecto de <c>Form1</c>— puede
/// cerrar en verde sin validar (W-1 corre siempre). Sustituye a la composición duplicada que dejaba
/// el 8.º argumento en <c>null</c> y convertía el constructor por defecto en una trampa si cambiaba
/// el entrypoint.
/// </summary>
public static class ComposicionServicios
{
    /// <summary>
    /// Compone los tres servicios del host: el procesador single-ASE, el procesador de período (5
    /// ASE) y el localizador de archivos fuente. Ambas vías de arranque (Program.cs y el constructor
    /// por defecto de Form1) usan ESTA composición.
    /// </summary>
    public static (IProcesadorRemuneracion Remuneracion, IProcesadorPeriodo Periodo, ArchivoFuenteLocator Locator) Crear()
    {
        var lectorR1R2R4 = new ExcelDataReaderRecaudoReader();
        var hojaLeafReader = new ExcelDataReaderWorkbookLeafInputReader();
        var calculo = new CalculoRemuneracion();
        var validador = new ValidadorBasico();
        var writer = new OpenXmlPlantillaWriter();
        var locator = new ArchivoFuenteLocator();

        // HU-20 (G3): oráculo R10 del período (OBLIGATORIO) — valida el DetRetri calculado en ambas
        // quincenas. HU-13 (2.7): oráculo de LECTURA de validaciones cruzadas, también OBLIGATORIO en
        // la composición del host: sin él la clasificación nunca corre y un run cerraría en verde sin
        // validar.
        var detRetriR10Reader = new ExcelDataReaderDetRetriR10Reader();
        var validacionOracleReader = new ValidacionOracleReader();

        IProcesadorRemuneracion procesador = new ProcesadorRemuneracion(
            lectorR1R2R4,
            hojaLeafReader,
            calculo,
            validador,
            writer);

        IProcesadorPeriodo procesadorPeriodo = new ProcesadorPeriodo(
            lectorR1R2R4,
            hojaLeafReader,
            calculo,
            validador,
            writer,
            locator,
            detRetriR10Reader,
            validacionOracleReader,
            // Plan 42 (T3-UI, R-UI-WIRING, D-E): emisor del reporte durable de divergencias
            // declaradas. Sin él, un cierre con declarada reconocida FALLA cerrado (sin reporte no
            // hay ÉXITO) en vez de silenciarse.
            new ReporteDivergenciasDeclaradasWriter());

        return (procesador, procesadorPeriodo, locator);
    }
}
