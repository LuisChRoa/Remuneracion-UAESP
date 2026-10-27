using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (H2.1): adaptador de FIxtures históricos. Los tests que asertaban coordenadas congeladas
/// (<c>F75</c>, <c>E47</c>, <c>D73</c>) se re-anclan a la IDENTIDAD del destino sin borrar la
/// aserción: la celda congelada se traduce a su ROL con el mapa LEGACY (que se conserva justamente
/// para esto) y el valor se lee del rol que ahora produce el lector.
///
/// Este tipo vive SOLO en el proyecto de tests: ningún camino de producción lo consume.
/// </summary>
internal static class RolsQ2
{
    /// <summary>Rol de la subsección oportuna (R1) para una ocurrencia y una columna dadas.</summary>
    internal static string Oportuna(int ocurrencia, bool especiales = false) =>
        Rol(ResolvedorDestinoValor.NivelDestino.Subseccion, ResolvedorDestinoValor.SubseccionOportuna, ocurrencia,
            especiales ? ResolvedorDestinoValor.EncabezadoEspecialesR1 : ResolvedorDestinoValor.EncabezadoTotal);

    /// <summary>Rol de la subsección extemporánea (R1) para una ocurrencia y una columna dadas.</summary>
    internal static string Extemporanea(int ocurrencia, bool especiales = false) =>
        Rol(ResolvedorDestinoValor.NivelDestino.Subseccion, ResolvedorDestinoValor.SubseccionExtemporanea, ocurrencia,
            especiales ? ResolvedorDestinoValor.EncabezadoEspecialesR1 : ResolvedorDestinoValor.EncabezadoTotal);

    /// <summary>Rol del cierre de una sección del R2.</summary>
    internal static string SeccionR2(string etiqueta, bool especiales = false) =>
        Rol(ResolvedorDestinoValor.NivelDestino.Seccion, etiqueta, 1,
            especiales ? ResolvedorDestinoValor.EncabezadoEspecialesR2 : ResolvedorDestinoValor.EncabezadoTotal);

    /// <summary>Rol del gran total del bloque R4.</summary>
    internal static string GranTotalR4(bool especiales = false) =>
        Rol(ResolvedorDestinoValor.NivelDestino.GranTotal, string.Empty, 1,
            especiales ? ResolvedorDestinoValor.EncabezadoEspecialesR4 : ResolvedorDestinoValor.EncabezadoTotal);

    /// <summary>Traduce una celda R1 congelada del mapa LEGACY a su rol de destino.</summary>
    internal static string DeCeldaR1(int aseId, string celda)
    {
        foreach (var (legado, fuente) in WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(aseId))
        {
            if (!string.Equals(legado, celda, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            return fuente switch
            {
                WorkbookLeafCellMapQ2.R1Q2Fuente.Mes0 => Oportuna(1),
                WorkbookLeafCellMapQ2.R1Q2Fuente.Mes1 => Oportuna(2),
                WorkbookLeafCellMapQ2.R1Q2Fuente.Mes2 => Oportuna(3),
                WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes0 => Oportuna(1, especiales: true),
                WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes1 => Oportuna(2, especiales: true),
                WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes2 => Oportuna(3, especiales: true),
                WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic0 => Extemporanea(1),
                WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic1 => Extemporanea(2),
                WorkbookLeafCellMapQ2.R1Q2Fuente.LAplic0 => Extemporanea(1, especiales: true),
                _ => throw new ArgumentOutOfRangeException(nameof(celda), celda, "Rol R1 no reconocido."),
            };
        }

        throw new ArgumentOutOfRangeException(nameof(celda), celda, $"La celda '{celda}' no está en el mapa LEGACY R1 del ASE {aseId}.");
    }

    /// <summary>Traduce una celda R2 congelada del mapa LEGACY a su rol de destino.</summary>
    internal static string DeCeldaR2(int aseId, string celda)
    {
        var mapa = WorkbookLeafCellMapQ2.ObtenerR2Q2Editables(aseId);
        if (string.Equals(celda, mapa.Componente, StringComparison.OrdinalIgnoreCase))
        {
            return SeccionR2("Componente");
        }

        if (string.Equals(celda, mapa.SubsCont, StringComparison.OrdinalIgnoreCase))
        {
            return SeccionR2("Subs/Cont");
        }

        if (string.Equals(celda, mapa.Especiales, StringComparison.OrdinalIgnoreCase))
        {
            return SeccionR2("Componente", especiales: true);
        }

        throw new ArgumentOutOfRangeException(nameof(celda), celda, $"La celda '{celda}' no está en el mapa LEGACY R2 del ASE {aseId}.");
    }

    /// <summary>Traduce una celda R4 congelada del mapa LEGACY a su rol de destino.</summary>
    internal static string DeCeldaR4(int aseId, string celda)
    {
        var mapa = WorkbookLeafCellMapQ2.ObtenerR4Q2Editables(aseId);
        if (string.Equals(celda, mapa.Total, StringComparison.OrdinalIgnoreCase))
        {
            return GranTotalR4();
        }

        if (string.Equals(celda, mapa.P, StringComparison.OrdinalIgnoreCase))
        {
            return GranTotalR4(especiales: true);
        }

        throw new ArgumentOutOfRangeException(nameof(celda), celda, $"La celda '{celda}' no está en el mapa LEGACY R4 del ASE {aseId}.");
    }

    private static string Rol(
        ResolvedorDestinoValor.NivelDestino nivel,
        string etiqueta,
        int ocurrencia,
        string encabezado) =>
        new ResolvedorDestinoValor.RolValor(nivel, etiqueta, ocurrencia, encabezado).Clave;
}
