using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 59 (H5, §0.1/§0.2): resuelve el nombre REAL de las hojas cuyo nombre lleva el sufijo del
/// período (<c>DetRetri</c> / <c>DetValiRetri</c>) leyendo la PLANTILLA, no una constante.
///
/// El sufijo se DERIVA del propio workbook: se enumeran las hojas, se identifica la raíz por su
/// identidad estructural (<c>DetRetri</c> + 7 dígitos) y se extrae su sufijo. Ese sufijo es el ancla
/// que la plantilla declara (la v2/v3 declaran uno propio); el nombre resuelto se compone
/// sustituyéndolo por <see cref="Periodo.CodigoCompleto"/> del período en proceso. Si el workbook ya
/// fue preparado (Phase 0 renombra las tres hojas del período) el nombre resuelto es el que la hoja
/// declara — la sustitución es identidad.
///
/// Consecuencia: <b>cero literales de período en el código</b>. Un ancla escrita a mano (p. ej. la
/// foto de un período concreto) queda estructuralmente fuera de alcance: la única fuente del sufijo
/// es el workbook que se está procesando, y una base de otro período falla con mensaje accionable en
/// vez de escribir en la hoja equivocada.
/// </summary>
internal static class HojasPeriodo
{
    /// <summary>Raíz del nombre de la hoja de detalle (R10).</summary>
    internal const string RaizDetRetri = "DetRetri";

    /// <summary>Raíz del nombre de la hoja de validación del detalle.</summary>
    internal const string RaizDetValiRetri = "DetValiRetri";

    /// <summary>Largo del sufijo de período (AAAAMMQ).</summary>
    private const int LargoSufijo = 7;

    /// <summary>Hoja <c>DetRetri</c> del período pedido, leída del workbook.</summary>
    internal static string DetRetri(WorkbookPart parte, Periodo periodo) =>
        Resolver(parte, RaizDetRetri, periodo);

    /// <summary>Hoja <c>DetValiRetri</c> del período pedido, leída del workbook.</summary>
    internal static string DetValiRetri(WorkbookPart parte, Periodo periodo) =>
        Resolver(parte, RaizDetValiRetri, periodo);

    private static string Resolver(WorkbookPart parte, string raiz, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(periodo);

        var hojas = (parte.Workbook?.Sheets?.Elements<Sheet>() ?? [])
            .Select(hoja => hoja.Name?.Value ?? string.Empty)
            .Where(nombre => nombre.Length > 0)
            .ToList();

        var declaradas = hojas.Where(nombre => EsHojaDeLaRaiz(nombre, raiz)).ToList();

        if (declaradas.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la plantilla declara {declaradas.Count} hojas '{raiz}' con sufijo de período y el período en proceso ('{periodo.CodigoCompleto}') exige exactamente una. " +
                $"Hojas halladas: {(declaradas.Count == 0 ? "(ninguna)" : string.Join(", ", declaradas))}. " +
                "El nombre se deriva de la plantilla; no se compone con un literal.");
        }

        var declarada = declaradas[0];

        // Ancla = sufijo que la plantilla declara HOY (leído, nunca escrito en el código).
        var ancla = declarada[raiz.Length..];

        // Sustitución del ancla por el código completo del período: el nombre de la hoja del período
        // en proceso. En una base ya preparada (Phase 0) la sustitución es identidad.
        var esperada = raiz + periodo.CodigoCompleto;

        if (string.Equals(declarada, esperada, StringComparison.Ordinal))
        {
            return declarada;
        }

        if (hojas.Exists(nombre => string.Equals(nombre, esperada, StringComparison.Ordinal)))
        {
            return esperada;
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-PLANTILLA: la hoja de detalle declara el sufijo '{ancla}' ('{declarada}') y el período en proceso es '{periodo.CodigoCompleto}' ('{esperada}' no existe en el workbook). " +
            "La base del período debe prepararse (Phase 0 renombra las hojas del período) o la plantilla debe corresponder al período: la app no escribe en la hoja de otro período.");
    }

    /// <summary>
    /// Identidad estructural de la hoja: raíz + sufijo numérico de período. No compara contra ningún
    /// valor concreto de período: la raíz es vocabulario de la plantilla y el sufijo se LEE.
    /// </summary>
    private static bool EsHojaDeLaRaiz(string nombre, string raiz) =>
        nombre.Length == raiz.Length + LargoSufijo
        && nombre.StartsWith(raiz, StringComparison.OrdinalIgnoreCase)
        && nombre[raiz.Length..].All(char.IsAsciiDigit);
}
