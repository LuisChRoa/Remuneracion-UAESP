using System.Globalization;
using System.Text;

namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 34 (T2, D-A/D-B — R-B-1/R-B-5): tabla de SINONIMIA de empresas del INTERIOR de
/// <c>Reporte Componentes R1</c> como DATOS citados, más el matcher puro
/// <see cref="SonMismaEmpresa"/>.
///
/// Motivo (Plan 33-T0, H1 CONFIRMADA): el template canónico trae el rótulo legado
/// <c>RECIPROCIDAD</c> (marca histórica de EAAB) mientras la fuente R1 trae el nombre
/// vigente <c>NUEVO ESQUEMA</c>; son la MISMA EFC (EAAB Reciprocidad), por lo que el filtro por
/// nombre del interior (<see cref="R1FirmaInterior.EsDatoEmpresa(FilaEspejoR1, string)"/>) debe
/// reconciliarlos.
///
/// Doctrina (Plan 32 D-B — catálogo C ABIERTO): el sistema NO conoce empresas por ramas de código.
/// Una divergencia futura legado↔vigente de CUALQUIER empresa se resuelve añadiendo una FILA citada
/// a <see cref="Clases"/>; el matcher es puro (strings → bool, sin I/O, sin OpenXML, sin períodos).
/// Cada fila exige su CITA de negocio (sin cita no hay fila — mitigación R-SINONIMO-ABUSO).
///
/// El matcher no recibe ningún período (restricción D-H: cero-quema-de-períodos). Un rótulo no
/// tabulado solo se iguala a sí mismo (catálogo abierto por construcción): las empresas sin
/// divergencia quedan identidad por construcción.
/// </summary>
public static class SinonimosEmpresaR1
{
    /// <summary>
    /// Clase de equivalencia: conjunto de rótulos que designan la MISMA empresa, con su cita de
    /// negocio (fuente que prueba la equivalencia). Los miembros se guardan normalizados
    /// (mayúsculas, sin diacríticos) para el match directo contra <see cref="SonMismaEmpresa"/>.
    /// </summary>
    public sealed record ClaseSinonimo(IReadOnlySet<string> Miembros, string Cita);

    /// <summary>
    /// Tabla de clases de equivalencia. Fila inicial única: EAAB Reciprocidad.
    /// </summary>
    public static IReadOnlyList<ClaseSinonimo> Clases { get; } =
    [
        new(
            new HashSet<string>(StringComparer.Ordinal) { "RECIPROCIDAD", "NUEVO ESQUEMA" },
            "Detalle de plantilla (REPORTE RECAUDO x BANCO): «NUEVO ESQUEMA: Corresponde al recaudo " +
            "\"EAAB Reciprocidad\", reportado por la empresa de facturación conjunta EAAB…»; " +
            "Proceso de Recaudo (matriz EFC×ASE): ASE2 = ENEL + EAAB (RECIPROCIDAD = marca histórica de EAAB); " +
            "oráculo del administrativo: Reporte Componentes R1!F199 = F140+F114-L114 (filas NUEVO ESQUEMA).")
    ];

    /// <summary>
    /// Verdad si <paramref name="a"/> y <paramref name="b"/> designan la misma empresa: iguales tras
    /// normalizar (mayúsculas, sin diacríticos) o miembros de una misma <see cref="ClaseSinonimo"/>.
    /// Rótulos vacíos o no tabulados solo se igualan a sí mismos (catálogo abierto).
    /// </summary>
    public static bool SonMismaEmpresa(string? a, string? b)
    {
        var normalizadaA = Normalizar(a);
        var normalizadaB = Normalizar(b);
        if (normalizadaA.Length == 0 || normalizadaB.Length == 0)
        {
            return false;
        }

        if (string.Equals(normalizadaA, normalizadaB, StringComparison.Ordinal))
        {
            return true;
        }

        foreach (var clase in Clases)
        {
            if (clase.Miembros.Contains(normalizadaA) && clase.Miembros.Contains(normalizadaB))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Normalización de rótulos de empresa: mayúsculas invariantes y sin diacríticos (misma regla que
    /// <c>Normalizar</c> del mutador del espejo / las firmas de <see cref="FilaEspejoR1"/>).
    /// </summary>
    private static string Normalizar(string? texto)
    {
        if (string.IsNullOrWhiteSpace(texto))
        {
            return string.Empty;
        }

        var descompuesto = texto.Normalize(NormalizationForm.FormD);
        var constructor = new StringBuilder(descompuesto.Length);
        foreach (var caracter in descompuesto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(caracter) != UnicodeCategory.NonSpacingMark)
            {
                constructor.Append(char.ToUpperInvariant(caracter));
            }
        }

        return constructor.ToString();
    }
}
