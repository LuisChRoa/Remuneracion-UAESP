using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// Orquestador del caso de uso de remuneración para un único ASE.
/// </summary>
public sealed class ProcesadorRemuneracion : IProcesadorRemuneracion
{
    private readonly IRecaudoReader _recaudoReader;
    private readonly IWorkbookLeafInputReader _leafReader;
    private readonly ICalculoRemuneracion _calculoRemuneracion;
    private readonly IValidador _validador;
    private readonly IWorkbookLeafWriter _workbookLeafWriter;

    public ProcesadorRemuneracion(
        IRecaudoReader recaudoReader,
        IWorkbookLeafInputReader leafReader,
        ICalculoRemuneracion calculoRemuneracion,
        IValidador validador,
        IWorkbookLeafWriter workbookLeafWriter)
    {
        _recaudoReader = recaudoReader ?? throw new ArgumentNullException(nameof(recaudoReader));
        _leafReader = leafReader ?? throw new ArgumentNullException(nameof(leafReader));
        _calculoRemuneracion = calculoRemuneracion ?? throw new ArgumentNullException(nameof(calculoRemuneracion));
        _validador = validador ?? throw new ArgumentNullException(nameof(validador));
        _workbookLeafWriter = workbookLeafWriter ?? throw new ArgumentNullException(nameof(workbookLeafWriter));
    }

    public ResultadoProcesoAse Ejecutar(SolicitudProcesoAse solicitud, IProgress<string>? progreso = null)
    {
        ArgumentNullException.ThrowIfNull(solicitud);

        progreso?.Report("Iniciando ejecución del procesador real.");
        progreso?.Report($"Periodo: {solicitud.Periodo.CodigoCompleto}; ASE: {solicitud.Ase.Id} - {solicitud.Ase.NombreCompleto}");

        if (string.IsNullOrWhiteSpace(solicitud.RutaR1) || string.IsNullOrWhiteSpace(solicitud.RutaR2) || string.IsNullOrWhiteSpace(solicitud.RutaR4))
        {
            throw new ArchivoFuenteNoEncontradoException("Debe indicarse R1, R2 y R4 para ejecutar el proceso.");
        }

        if (string.IsNullOrWhiteSpace(solicitud.RutaPlantilla) || string.IsNullOrWhiteSpace(solicitud.RutaSalida))
        {
            throw new ArchivoFuenteNoEncontradoException("Debe indicarse plantilla y ruta de salida.");
        }

        progreso?.Report("Leyendo R1, R2 y R4...");
        var r1 = _recaudoReader.LeerR1(solicitud.RutaR1);
        var r2 = _recaudoReader.LeerR2(solicitud.RutaR2);
        var r4 = _recaudoReader.LeerR4(solicitud.RutaR4);

        progreso?.Report("Calculando consolidado del ASE...");
        var resultado = _calculoRemuneracion.CalcularConsolidado(solicitud.Periodo, [(solicitud.Ase, r1, r2, r4)]);

        progreso?.Report("Leyendo inputs leaf del workbook...");
        var leaf = _leafReader.LeerLeafInputs(solicitud.Ase, solicitud.Periodo, solicitud.RutaR1, solicitud.RutaR2, solicitud.RutaR4);

        progreso?.Report("Validando coherencia básica...");
        var errores = _validador.Validar(resultado, leaf);
        if (errores.Count > 0)
        {
            var detalle = string.Join("; ", errores);
            throw new CalculoInvalidoException($"La validación básica falló: {detalle}");
        }

        progreso?.Report("Generando workbook de salida...");
        _workbookLeafWriter.GenerarWorkbook(solicitud.RutaPlantilla, solicitud.RutaSalida, resultado, leaf);

        progreso?.Report("Proceso completado correctamente.");

        return new ResultadoProcesoAse
        {
            Resultado = resultado,
            Leaf = leaf,
            RutaSalida = solicitud.RutaSalida
        };
    }
}
