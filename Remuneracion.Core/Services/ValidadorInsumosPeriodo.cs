using Remuneracion.Core.Constants;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// Plan 26 (T1, §2.1/D-D/D-F): preflight de insumos del período. Enumera la lista COMPLETA de
/// insumos faltantes ANTES de abrir ningún workbook, reutilizando los MISMOS finders del runtime
/// vía <see cref="ILocalizadorArchivosAse"/> (D-C: cero lógica de match propia, cero divergencia).
/// Solo verifica EXISTENCIA en filesystem; no valida contenido (non-goal §1.3) y no abre workbooks.
/// <c>RecaudosReversados</c> queda FUERA: el runtime no lo consume (§V9).
/// </summary>
public sealed class ValidadorInsumosPeriodo
{
    private readonly ILocalizadorArchivosAse _localizador;

    /// <summary>
    /// Crea el preflight sobre el localizador existente (misma abstracción que usan los
    /// procesadores; DIP intacto).
    /// </summary>
    /// <param name="localizador">Contrato de localización de carpetas/archivos del período.</param>
    public ValidadorInsumosPeriodo(ILocalizadorArchivosAse localizador)
    {
        _localizador = localizador ?? throw new ArgumentNullException(nameof(localizador));
    }

    /// <summary>
    /// Devuelve la lista completa y ordenada de insumos faltantes del período (vacía si no falta
    /// nada). Orden estable (D-F): ASE 1..5 (carpeta, luego reportes), conciliaciones, R10.
    /// </summary>
    /// <param name="carpetaPeriodo">Carpeta del período que contiene las 5 carpetas de ASE.</param>
    /// <param name="periodo">Período; su quincena gobierna la exigencia de saldos-notas/retribución.</param>
    public IReadOnlyList<InsumoFaltante> Validar(string carpetaPeriodo, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(periodo);

        var faltantes = new List<InsumoFaltante>();
        if (string.IsNullOrWhiteSpace(carpetaPeriodo))
        {
            // El guard del procesador ya nombra la carpeta faltante; el preflight no aporta más.
            return faltantes;
        }

        var esQuincena2 = periodo.NumeroQuincena == 2;
        var carpetasAse = _localizador.ObtenerCarpetasAse(carpetaPeriodo);

        for (var aseId = 1; aseId <= CarpetasAse.Prefijos.Count; aseId++)
        {
            var carpetaAse = carpetasAse.FirstOrDefault(c =>
                Path.GetFileName(c).StartsWith(aseId.ToString(), StringComparison.OrdinalIgnoreCase));

            if (carpetaAse is null)
            {
                // D-F: carpeta ausente → un ítem que la nombra + sus reportes; se omiten sus archivos.
                faltantes.Add(FormateadorInsumosFaltantes.CarpetaAse(aseId, periodo));
                continue;
            }

            // Mismos finders que el loop runtime de ProcesadorPeriodo (mismos prefijos/fallback).
            if (_localizador.BuscarArchivo(carpetaAse, "Recaudoporcomponente") is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.R1));
            }

            if (_localizador.BuscarArchivo(carpetaAse, "RerpoteDetalleSaldosaFavor") is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.R2));
            }

            if ((_localizador.BuscarArchivo(carpetaAse, "ReversiónPorComponente")
                 ?? _localizador.BuscarArchivo(carpetaAse, "ReversionPorComponente")) is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.R4));
            }

            if (_localizador.BuscarReporteBanco(carpetaAse) is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.Banco));
            }

            if (_localizador.BuscarBalance(carpetaAse) is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.Balance));
            }

            if (esQuincena2)
            {
                // D-D: saldos-notas y retribución-negativa SOLO en Q2 (gobierno por dominio).
                if (_localizador.BuscarSaldosNotas(carpetaAse) is null)
                {
                    faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.SaldosNotas));
                }

                if (_localizador.BuscarRetribucionNegativa(carpetaAse) is null)
                {
                    faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.RetribucionNegativa));
                }
            }
        }

        // Insumos de período: 5 conciliaciones SIEMPRE (Q1 y Q2, §V3) + R10 SIEMPRE.
        AgregarConciliaciones(carpetaPeriodo, faltantes);

        if (_localizador.BuscarR10(carpetaPeriodo) is null)
        {
            faltantes.Add(FormateadorInsumosFaltantes.R10(periodo, carpetaPeriodo));
        }

        return faltantes;
    }

    /// <summary>
    /// Evalúa las 5 conciliaciones por empresa. D-F: si la carpeta <c>Conciliaciones</c> completa
    /// está ausente, UN solo ítem agrega la carpeta + sus 5 archivos; si existe, un ítem por
    /// empresa faltante.
    /// </summary>
    private void AgregarConciliaciones(string carpetaPeriodo, List<InsumoFaltante> faltantes)
    {
        var carpetaConciliaciones = Path.Combine(carpetaPeriodo, "Conciliaciones");
        if (!Directory.Exists(carpetaConciliaciones))
        {
            faltantes.Add(FormateadorInsumosFaltantes.CarpetaConciliaciones());
            return;
        }

        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            if (_localizador.BuscarConciliacion(carpetaPeriodo, empresa.PrefijoConciliacion) is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ConciliacionEmpresa(empresa));
            }
        }
    }
}
