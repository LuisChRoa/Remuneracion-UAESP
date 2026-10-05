using System.Text;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// Plan 26 (T1, §2.1/D-E): SITIO ÚNICO que compone el texto del preflight de insumos en lenguaje
/// administrativo. <see cref="ValidadorInsumosPeriodo"/> solo enumera hechos (qué finder devolvió
/// <c>null</c>) invocando estas factorías; acá viven los nombres de reportes, las rutas esperadas
/// y las acciones. Prohibida la jerga técnica (nombres de clase, "finder", "runtime" o la palabra
/// "prefijo"): el archivo esperado se nombra como "su nombre debe comenzar con ...".
/// </summary>
public static class FormateadorInsumosFaltantes
{
    /// <summary>
    /// Encabezado + lista numerada COMPLETA de faltantes del período (5 ASE). El detalle numerado
    /// viaja en <c>ex.Message</c> y la guía existente (<c>CatalogoErrores.Para</c>) lo envuelve sin
    /// cambios (D-B).
    /// </summary>
    public static string Mensaje(Periodo periodo, IReadOnlyList<InsumoFaltante> faltantes)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(faltantes);

        var encabezado =
            $"Faltan insumos para el período {periodo.CodigoCompleto} (quincena {periodo.NumeroQuincena}). " +
            "No se procesó ningún ASE.";
        return Componer(encabezado, faltantes);
    }

    /// <summary>
    /// Encabezado + lista numerada de faltantes del flujo single-ASE (R1/R2/R4). Mismo código y
    /// mismo formateador que el período (R-F-5); solo cambia el alcance del encabezado.
    /// </summary>
    public static string MensajeAse(Periodo periodo, Ase ase, IReadOnlyList<InsumoFaltante> faltantes)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(faltantes);

        var encabezado =
            $"Faltan insumos para el período {periodo.CodigoCompleto} (quincena {periodo.NumeroQuincena}). " +
            $"No se procesó el ASE {ase.Id}.";
        return Componer(encabezado, faltantes);
    }

    /// <summary>
    /// D-F: carpeta de ASE ausente → UN ítem la nombra y enumera los reportes que debería contener
    /// (no una línea por archivo). El nombre de carpeta esperado es el prefijo de dominio.
    /// </summary>
    internal static InsumoFaltante CarpetaAse(int aseId, Periodo periodo)
    {
        var ase = AseFactory.DesdeId(aseId);
        var nombreCarpeta = CarpetasAse.Prefijos[aseId - 1];
        return new InsumoFaltante
        {
            Alcance = $"ASE {aseId}",
            QueFalta = $"Falta la carpeta del ASE {aseId} ({ase.NombreCompleto}).",
            DondeDebeIr =
                $"Debe estar dentro de la carpeta del período con el nombre exacto \"{nombreCarpeta}\" y contener: " +
                $"{ListaReportesAse(periodo.NumeroQuincena == 2)}.",
            QueHacer = "Solicítela con sus reportes al área encargada, colóquela con ese nombre exacto y vuelva a ejecutar."
        };
    }

    /// <summary>
    /// Reporte de ASE ausente (D-E): nombra el ASE, la carpeta donde debe estar y el inicio de
    /// nombre de archivo esperado.
    /// </summary>
    internal static InsumoFaltante ReporteAse(int aseId, string carpetaAse, ReporteInsumoAse reporte)
    {
        var ase = AseFactory.DesdeId(aseId);
        var (nombre, inicioNombre) = DescribirReporte(reporte);
        var nombreCarpeta = Path.GetFileName(carpetaAse);
        return new InsumoFaltante
        {
            Alcance = $"ASE {aseId}",
            QueFalta = $"Falta {nombre} del ASE {aseId} ({ase.NombreCompleto}).",
            DondeDebeIr =
                $"El archivo debe estar en la carpeta \"{nombreCarpeta}\" (dentro de la carpeta del período) " +
                $"y su nombre debe comenzar con \"{inicioNombre}\".",
            QueHacer = "Solicítelo al área encargada, colóquelo en esa carpeta y vuelva a ejecutar."
        };
    }

    /// <summary>
    /// Guardrail single-ASE (R-F-5): reporte ausente en una ruta explícita. No hay carpeta de
    /// período resuelta, así que el "dónde" es la ruta indicada por el usuario.
    /// </summary>
    internal static InsumoFaltante ReporteAseRuta(int aseId, string ruta, ReporteInsumoAse reporte)
    {
        var ase = AseFactory.DesdeId(aseId);
        var (nombre, _) = DescribirReporte(reporte);
        return new InsumoFaltante
        {
            Alcance = $"ASE {aseId}",
            QueFalta = $"Falta {nombre} del ASE {aseId} ({ase.NombreCompleto}).",
            DondeDebeIr = $"El archivo debe estar en la ruta \"{ruta}\".",
            QueHacer = "Solicítelo al área encargada, colóquelo en esa ruta y vuelva a ejecutar."
        };
    }

    /// <summary>
    /// D-F: si falta la carpeta <c>Conciliaciones</c> completa, UN solo ítem la nombra y enumera
    /// los 5 archivos esperados (no cinco líneas redundantes).
    /// </summary>
    internal static InsumoFaltante CarpetaConciliaciones()
    {
        return new InsumoFaltante
        {
            Alcance = "Período",
            QueFalta = "Falta la carpeta \"Conciliaciones\" dentro de la carpeta del período.",
            DondeDebeIr = $"Debe contener los 5 archivos de conciliación: {ListaEmpresasConciliacion()}.",
            QueHacer = "Solicítelos al área encargada, cree la carpeta con el nombre exacto \"Conciliaciones\", colóquelos ahí y vuelva a ejecutar."
        };
    }

    /// <summary>
    /// S6: la carpeta <c>Conciliaciones</c> existe pero falta el archivo de una empresa → un ítem
    /// que la nombra con su nombre administrativo y el inicio de nombre esperado.
    /// </summary>
    internal static InsumoFaltante ConciliacionEmpresa(EmpresaFacturacion empresa)
    {
        return new InsumoFaltante
        {
            Alcance = "Período",
            QueFalta = $"Falta el archivo de conciliación de {NombreEmpresa(empresa)}.",
            DondeDebeIr =
                "El archivo debe estar en la carpeta \"Conciliaciones\" (dentro de la carpeta del período) " +
                $"y su nombre debe comenzar con \"{empresa.PrefijoConciliacion}\".",
            QueHacer = "Solicítelo al área encargada, colóquelo en esa carpeta y vuelva a ejecutar."
        };
    }

    /// <summary>
    /// Insumo de período R10 (oráculo obligatorio en ambas quincenas, HU-20/G3). El "dónde"
    /// incluye la carpeta del período completa para que el usuario la ubique sin ambigüedad.
    /// </summary>
    internal static InsumoFaltante R10(Periodo periodo, string carpetaPeriodo)
    {
        return new InsumoFaltante
        {
            Alcance = "Período",
            QueFalta = $"Falta el archivo R10_Remuneracion_{periodo.CodigoCompleto}.xlsx en la carpeta del período.",
            DondeDebeIr = $"El archivo debe estar en la carpeta \"{carpetaPeriodo}\".",
            QueHacer = "Solicítelo al área encargada, colóquelo en esa carpeta y vuelva a ejecutar."
        };
    }

    /// <summary>
    /// Nombre administrativo del reporte de ASE + inicio de nombre de archivo esperado (fuente
    /// única de los textos por reporte).
    /// </summary>
    private static (string Nombre, string InicioNombre) DescribirReporte(ReporteInsumoAse reporte) => reporte switch
    {
        ReporteInsumoAse.R1 => ("Recaudo por componente (R1)", "Recaudoporcomponente"),
        ReporteInsumoAse.R2 => ("Detalle de saldos a favor (R2)", "RerpoteDetalleSaldosaFavor"),
        ReporteInsumoAse.R4 => ("Reversión por componente (R4)", "ReversiónPorComponente"),
        ReporteInsumoAse.Banco => ("Reporte de recaudo por banco", "ReportePagosxBanco"),
        ReporteInsumoAse.Balance => ("Balance de subsidios y contribuciones", "R4-BalanceSubsidioyContribuciones"),
        ReporteInsumoAse.SaldosNotas => ("Saldos a favor aplicados por notas", "SaldosaFavorAplicadosPorNotas"),
        ReporteInsumoAse.RetribucionNegativa => ("Retribución negativa", "RetribuciónNegativa"),
        _ => throw new ArgumentOutOfRangeException(nameof(reporte), reporte, "Reporte de ASE desconocido.")
    };

    /// <summary>
    /// Lista de reportes de ASE según la quincena (D-D): saldos-notas y retribución-negativa solo
    /// en quincena 2.
    /// </summary>
    private static string ListaReportesAse(bool esQuincena2)
    {
        var reportes = new List<string>
        {
            "Recaudo por componente (R1)",
            "Detalle de saldos a favor (R2)",
            "Reversión por componente (R4)",
            "Reporte de recaudo por banco",
            "Balance de subsidios y contribuciones"
        };

        if (esQuincena2)
        {
            reportes.Add("Saldos a favor aplicados por notas");
            reportes.Add("Retribución negativa");
        }

        return Enumerar(reportes);
    }

    /// <summary>Lista de las 5 empresas de conciliación con sus nombres administrativos (D-F).</summary>
    private static string ListaEmpresasConciliacion() =>
        Enumerar(EmpresaFacturacion.Catalogo.Select(NombreEmpresa));

    /// <summary>Nombre administrativo de una empresa de facturación (catálogo 2.2).</summary>
    private static string NombreEmpresa(EmpresaFacturacion empresa) => empresa.Id switch
    {
        1 => "Reciprocidad EAAB",
        2 => "ENEL",
        3 => "ENERBIT",
        4 => "Occidente Directa",
        5 => "Otros",
        _ => empresa.Nombre
    };

    /// <summary>Une una lista con comas y "y" antes del último elemento (sin anglicismos).</summary>
    private static string Enumerar(IEnumerable<string> elementos)
    {
        var lista = elementos.ToList();
        return lista.Count switch
        {
            0 => string.Empty,
            1 => lista[0],
            _ => string.Join(", ", lista.Take(lista.Count - 1)) + " y " + lista[^1]
        };
    }

    private static string Componer(string encabezado, IReadOnlyList<InsumoFaltante> faltantes)
    {
        var sb = new StringBuilder();
        sb.Append(encabezado);
        for (var i = 0; i < faltantes.Count; i++)
        {
            var faltante = faltantes[i];
            sb.Append(Environment.NewLine);
            sb.Append(i + 1).Append(") ");
            sb.Append(faltante.QueFalta).Append(' ');
            sb.Append(faltante.DondeDebeIr).Append(' ');
            sb.Append(faltante.QueHacer);
        }

        return sb.ToString();
    }
}
