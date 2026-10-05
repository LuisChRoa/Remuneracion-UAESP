using Serilog;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// Plan 26 (T1, §2.1/D-D/D-F): preflight de insumos del período. Enumera la lista COMPLETA de
/// insumos faltantes o inválidos ANTES de abrir ningún workbook, reutilizando los MISMOS finders
/// del runtime vía <see cref="ILocalizadorArchivosAse"/> (D-C: cero lógica de match propia, cero
/// divergencia).
/// Plan 27 (T2, D-B): además de la EXISTENCIA, valida la FIRMA binaria de cada path resuelto
/// (<see cref="InspectorFirmaExcel"/>, ≤<see cref="InspectorFirmaExcel.BytesALeer"/> B por archivo).
/// No abre workbooks ni valida contenido/headers/estructura (siguen fallando en sus readers).
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
            // Plan 27 (D-B): cada path resuelto se firma-chequea; si el archivo existe pero no es
            // un Excel válido, ocupa su lugar en la lista con un ítem administrativo.
            var rutaR1 = _localizador.BuscarArchivo(carpetaAse, "Recaudoporcomponente");
            if (rutaR1 is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.R1));
            }
            else
            {
                AgregarSiFirmaInvalida(rutaR1, $"ASE {aseId}", DescripcionCarpetaAse(carpetaAse), faltantes);
            }

            var rutaR2 = _localizador.BuscarArchivo(carpetaAse, "RerpoteDetalleSaldosaFavor");
            if (rutaR2 is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.R2));
            }
            else
            {
                AgregarSiFirmaInvalida(rutaR2, $"ASE {aseId}", DescripcionCarpetaAse(carpetaAse), faltantes);
            }

            var rutaR4 = _localizador.BuscarArchivo(carpetaAse, "ReversiónPorComponente")
                ?? _localizador.BuscarArchivo(carpetaAse, "ReversionPorComponente");
            if (rutaR4 is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.R4));
            }
            else
            {
                AgregarSiFirmaInvalida(rutaR4, $"ASE {aseId}", DescripcionCarpetaAse(carpetaAse), faltantes);
            }

            var rutaBanco = _localizador.BuscarReporteBanco(carpetaAse);
            if (rutaBanco is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.Banco));
            }
            else
            {
                AgregarSiFirmaInvalida(rutaBanco, $"ASE {aseId}", DescripcionCarpetaAse(carpetaAse), faltantes);
            }

            var rutaBalance = _localizador.BuscarBalance(carpetaAse);
            if (rutaBalance is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.Balance));
            }
            else
            {
                AgregarSiFirmaInvalida(rutaBalance, $"ASE {aseId}", DescripcionCarpetaAse(carpetaAse), faltantes);
            }

            if (esQuincena2)
            {
                // D-D: saldos-notas y retribución-negativa SOLO en Q2 (gobierno por dominio).
                var rutaSaldosNotas = _localizador.BuscarSaldosNotas(carpetaAse);
                if (rutaSaldosNotas is null)
                {
                    faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.SaldosNotas));
                }
                else
                {
                    AgregarSiFirmaInvalida(rutaSaldosNotas, $"ASE {aseId}", DescripcionCarpetaAse(carpetaAse), faltantes);
                }

                var rutaRetribucionNegativa = _localizador.BuscarRetribucionNegativa(carpetaAse);
                if (rutaRetribucionNegativa is null)
                {
                    faltantes.Add(FormateadorInsumosFaltantes.ReporteAse(aseId, carpetaAse, ReporteInsumoAse.RetribucionNegativa));
                }
                else
                {
                    AgregarSiFirmaInvalida(rutaRetribucionNegativa, $"ASE {aseId}", DescripcionCarpetaAse(carpetaAse), faltantes);
                }
            }
        }

        // Insumos de período: 5 conciliaciones SIEMPRE (Q1 y Q2, §V3) + R10 SIEMPRE.
        AgregarConciliaciones(carpetaPeriodo, faltantes);

        var rutaR10 = _localizador.BuscarR10(carpetaPeriodo);
        if (rutaR10 is null)
        {
            faltantes.Add(FormateadorInsumosFaltantes.R10(periodo, carpetaPeriodo));
        }
        else
        {
            AgregarSiFirmaInvalida(rutaR10, "Período", $"la carpeta \"{carpetaPeriodo}\"", faltantes);
        }

        return faltantes;
    }

    /// <summary>
    /// Plan 27 (T2, D-B): valida la firma de un path resuelto. Si es válida no hace nada; si no,
    /// registra el detalle técnico (firma + bytes) SOLO en el log y agrega el ítem administrativo
    /// correspondiente (variante web o formato desconocido; <see cref="FirmaExcel.Ilegible"/> se
    /// trata como desconocido, D-C).
    /// </summary>
    private static void AgregarSiFirmaInvalida(
        string rutaArchivo,
        string alcance,
        string descripcionUbicacion,
        List<InsumoFaltante> faltantes)
    {
        var resultado = InspectorFirmaExcel.Inspeccionar(rutaArchivo);
        if (resultado.Firma == FirmaExcel.Valida)
        {
            return;
        }

        // Jerga técnica (firma detectada + bytes) SOLO al log; el mensaje al usuario es administrativo.
        Log.Warning(
            "Preflight: el archivo {Archivo} no tiene firma de Excel ({Firma}); primeros bytes {PrimerosBytes}.",
            rutaArchivo,
            resultado.Firma,
            resultado.PrimerosBytes.Length == 0 ? "(vacío)" : Convert.ToHexString(resultado.PrimerosBytes));

        faltantes.Add(resultado.Firma == FirmaExcel.NoExcelPareceCopiaWeb
            ? FormateadorInsumosFaltantes.ArchivoNoEsExcelPareceWeb(alcance, rutaArchivo, descripcionUbicacion)
            : FormateadorInsumosFaltantes.ArchivoNoEsExcelFormatoDesconocido(alcance, rutaArchivo, descripcionUbicacion));
    }

    /// <summary>Ubicación administrativa esperada de un reporte de ASE (D-E).</summary>
    private static string DescripcionCarpetaAse(string carpetaAse) =>
        $"la carpeta \"{Path.GetFileName(carpetaAse)}\" (dentro de la carpeta del período)";

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
            var ruta = _localizador.BuscarConciliacion(carpetaPeriodo, empresa.PrefijoConciliacion);
            if (ruta is null)
            {
                faltantes.Add(FormateadorInsumosFaltantes.ConciliacionEmpresa(empresa));
            }
            else
            {
                // Plan 27 (D-B): la conciliación resuelta también pasa por firma-chequeo.
                AgregarSiFirmaInvalida(ruta, "Período", "la carpeta \"Conciliaciones\" (dentro de la carpeta del período)", faltantes);
            }
        }
    }
}
