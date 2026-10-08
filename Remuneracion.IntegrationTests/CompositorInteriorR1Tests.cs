using System.Globalization;
using System.Text.RegularExpressions;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 32 (T2-interior, RONDA A — R-B-1): tests unitarios del compositor PURO
/// <see cref="CompositorInteriorR1"/> contra el fixture congelado del manual
/// (<see cref="InterioresR1Esperados.Agosto"/>).
///
/// Sin workbook, sin Excel, sin I/O: por cada entrada R-1 de agosto se construye <c>filasPorFirma</c>
/// con las filas REALES extraídas del propio texto esperado del manual y se verifica que el compositor
/// reproduce EXACTO ese texto. Esto prueba la REGLA DE COMPOSICIÓN (signos, orden y cardinalidad
/// 2↔3 Mes) sin abrir un archivo. Para las entradas clase L-1 (<c>Formula</c> nulo, el manual no trae
/// <c>&lt;f&gt;</c>) se aserta <c>null</c>. Los casos de hoja se resuelven por firma: el espejo G/H/L se
/// hermana con su subtotal-empresa F del mismo renglón y el EXTEMP G/H con su EXTEMP F del mismo renglón.
/// </summary>
public sealed class CompositorInteriorR1Tests
{
    private static readonly Regex Termino = new(
        @"(?<signo>[+-]?)(?<col>[A-Za-z]+)(?<fila>\d+)",
        RegexOptions.Compiled);

    public static IEnumerable<object[]> Agosto =>
        InterioresR1Esperados.Agosto.Select(v => new object[] { v.AseId, v.Celda, v.Etiqueta, v.Formula! });

    [Theory]
    [MemberData(nameof(Agosto))]
    public void ComponerInterior_Agosto_ReproduceTextoDelManual(int aseId, string celda, string etiqueta, string? formula)
    {
        if (formula is null)
        {
            // Clase L-1: el manual no trae <f>. Sin anclas el compositor no inventa filas -> null.
            Assert.Null(CompositorInteriorR1.ComponerInterior(aseId, celda, new Dictionary<string, int>(), null));
            return;
        }

        var filasPorFirma = FilasPorFirma(aseId, celda, etiqueta);
        var texto = CompositorInteriorR1.ComponerInterior(aseId, celda, filasPorFirma, null);

        Assert.Equal(formula, texto);
    }

    [Fact]
    public void SinAnclas_DevuelveNull_YAnclasFaltantesLoReporta()
    {
        var vacio = new Dictionary<string, int>();

        Assert.Null(CompositorInteriorR1.ComponerInterior(1, "F60", vacio, null));
        Assert.Contains(CompositorInteriorR1.EtiquetaMes + "0", CompositorInteriorR1.AnclasFaltantes("SUB_EMP", vacio, null));
        Assert.Contains(CompositorInteriorR1.EtiquetaAncla, CompositorInteriorR1.AnclasFaltantes("SUBS", vacio, null));
    }

    [Fact]
    public void ExtIntCeroAplic_DevuelveNull_YAnclasFaltantesNombraAplic0()
    {
        var vacio = new Dictionary<string, int>();
        var inexistentes = new HashSet<string>(StringComparer.Ordinal) { CompositorInteriorR1.EtiquetaAplic + "0" };

        Assert.Null(CompositorInteriorR1.ComponerInterior(3, "F327", vacio, inexistentes));
        Assert.Contains(
            CompositorInteriorR1.EtiquetaAplic + "0",
            CompositorInteriorR1.AnclasFaltantes("EXT_INT", vacio, inexistentes));
    }

    [Fact]
    public void QuirkDobleSignoL_SeReproduceTalCual()
    {
        // F204 del manual: F133+F105+F87--L87-L105 (H-DOBLE-SIGNO reproducido, nunca normalizado).
        var filas = new Dictionary<string, int>(StringComparer.Ordinal)
        {
            ["Mes0"] = 133,
            ["Mes1"] = 105,
            ["Mes2"] = 87,
            [CompositorInteriorR1.EtiquetaQuirkDobleSignoL] = 0
        };

        Assert.Equal("F133+F105+F87--L87-L105", CompositorInteriorR1.ComponerInterior(2, "F204", filas, null));

        // Sin el marcador, el mismo bloque compone el signo simple.
        filas.Remove(CompositorInteriorR1.EtiquetaQuirkDobleSignoL);
        Assert.Equal("F133+F105+F87-L87-L105", CompositorInteriorR1.ComponerInterior(2, "F204", filas, null));
    }

    private static IReadOnlyDictionary<string, int> FilasPorFirma(int aseId, string celda, string etiqueta)
    {
        var filas = new Dictionary<string, int>(StringComparer.Ordinal);

        switch (etiqueta)
        {
            case "SUB_EMP":
            {
                var entrada = Buscar(aseId, celda);
                AsignarPrefijo(filas, CompositorInteriorR1.EtiquetaMes, TokensDeColumna(entrada.Formula!, "F"));
                if (entrada.Formula!.Contains("--L", StringComparison.Ordinal))
                {
                    filas[CompositorInteriorR1.EtiquetaQuirkDobleSignoL] = 0;
                }

                break;
            }

            case "SUB_TDF":
            case "SUB_L":
            {
                var hermano = InterioresR1Esperados.Agosto.First(v =>
                    v.AseId == aseId
                    && v.Etiqueta == "SUB_EMP"
                    && Fila(v.Celda) == Fila(celda));
                AsignarPrefijo(filas, CompositorInteriorR1.EtiquetaMes, TokensDeColumna(hermano.Formula!, "F"));
                break;
            }

            case "EXT_INT":
            {
                var origen = Columna(celda) is "G" or "H"
                    ? InterioresR1Esperados.Agosto.First(v =>
                        v.AseId == aseId
                        && v.Etiqueta == "EXT_INT"
                        && Columna(v.Celda) == "F"
                        && Fila(v.Celda) == Fila(celda))
                    : Buscar(aseId, celda);
                AsignarPrefijo(filas, CompositorInteriorR1.EtiquetaAplic, TokensDeColumna(origen.Formula!, "F"));
                break;
            }

            case "SUBS":
            case "AFASEO":
            {
                var entrada = Buscar(aseId, celda);
                var token = Termino.Matches(entrada.Formula!)[0];
                filas[CompositorInteriorR1.EtiquetaAncla] =
                    int.Parse(token.Groups["fila"].Value, CultureInfo.InvariantCulture);
                break;
            }

            default:
                throw new InvalidOperationException($"Etiqueta interior no contemplada: '{etiqueta}'.");
        }

        return filas;
    }

    private static void AsignarPrefijo(Dictionary<string, int> filas, string prefijo, IReadOnlyList<int> valores)
    {
        for (var i = 0; i < valores.Count; i++)
        {
            filas[prefijo + i.ToString(CultureInfo.InvariantCulture)] = valores[i];
        }
    }

    private static IReadOnlyList<int> TokensDeColumna(string formula, string columna) =>
        Termino.Matches(formula)
            .Where(m => string.Equals(m.Groups["col"].Value, columna, StringComparison.OrdinalIgnoreCase))
            .Select(m => int.Parse(m.Groups["fila"].Value, CultureInfo.InvariantCulture))
            .ToList();

    private static InterioresR1Esperados.InteriorR1 Buscar(int aseId, string celda) =>
        InterioresR1Esperados.Agosto.First(v => v.AseId == aseId && string.Equals(v.Celda, celda, StringComparison.Ordinal));

    private static string Columna(string celda)
    {
        var letras = new string(celda.Where(char.IsLetter).ToArray());
        return letras.ToUpperInvariant();
    }

    private static int Fila(string celda) =>
        int.Parse(new string(celda.Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture);
}
