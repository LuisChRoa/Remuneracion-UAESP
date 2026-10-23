using System.Text.RegularExpressions;
using Remuneracion.Core.Models;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 55 (H1.2/H1.5 + D5, CA7/G-D-H): CANDADO ANTI-QUEMADO del camino de resolución de la
/// conciliación por empresa.
///
/// Patrón del candado del Plan 54 (<c>LecturaR1SinRamaQuincenaTests</c>): test ESTÁTICO con
/// <b>allow-list cerrada</b> y <b>prueba de injerto</b> (una lista por ASE o un literal de empresa
/// inyectado a propósito debe enrojecerlo). Lo que el candado prohíbe en el camino de resolución:
/// <list type="bullet">
/// <item>listas / ramas por ASE (<c>aseId switch</c>, <c>case 1:</c>, <c>Ase.Id ==</c>);</item>
/// <item>literales de empresa (los rótulos del catálogo y los alias de la sinonimia);</item>
/// <item>la tabla congelada por (empresa, ASE) (<c>WorkbookLeafCellMapPorEmpresa</c>, LEGACY).</item>
/// </list>
/// La doctrina manda lo contrario: la plantilla es la autoridad (cruce de etiquetas) y el vocabulario
/// de empresas sale del catálogo + <see cref="SinonimosEmpresaR1"/> (catálogo abierto).
/// </summary>
public sealed class ConciliacionEmpresaCandadoAntiQuemadoTests
{
    /// <summary>Archivos del CAMINO DE RESOLUCIÓN (allow-list cerrada): resolvedor + escritura por empresa.</summary>
    private static readonly string[] CaminoDeResolucion =
    [
        Path.Combine("Remuneracion.Infrastructure", "Excel", "ResolvedorConciliacionEmpresa.cs"),
    ];

    /// <summary>Métodos del writer que consumen la resolución (se escanean por región acotada).</summary>
    private static readonly (string Archivo, string Desde, string Hasta)[] RegionesDelWriter =
    [
        (
            Path.Combine("Remuneracion.Infrastructure", "Excel", "OpenXmlPlantillaWriter.cs"),
            "private static void EscribirCeldasConciliacion(",
            "private static void EscribirCeldasBanco(")
    ];

    /// <summary>
    /// Rótulos de empresa que el camino de resolución NO puede escribir como literal (catálogo 2.2 +
    /// alias citado de la sinonimia + la etiqueta real que declara el cuadro de ASE2).
    /// </summary>
    private static readonly string[] RotulosDeEmpresa =
    [
        "ENEL", "OCCIDENTE", "ENERBIT", "RECIPROCIDAD", "NUEVO ESQUEMA", "CIUDAD LIMPIA-ACUEDUCTO", "EAB"
    ];

    [Fact]
    public void T4_CaminoDeResolucion_SinListasPorAseNiLiteralesDeEmpresa()
    {
        foreach (var relativo in CaminoDeResolucion)
        {
            var fuente = File.ReadAllText(Ruta(relativo));
            var violaciones = Analizar(fuente);

            Assert.True(
                violaciones.Count == 0,
                $"{relativo} viola G-D-H/CA7 (listas por ASE o literales de empresa en el camino de resolución):"
                + Environment.NewLine + string.Join(Environment.NewLine, violaciones));
        }

        foreach (var (archivo, desde, hasta) in RegionesDelWriter)
        {
            var fuente = Region(Ruta(archivo), desde, hasta);
            var violaciones = Analizar(fuente);

            Assert.True(
                violaciones.Count == 0,
                $"{archivo} [{desde} .. {hasta}] viola G-D-H/CA7:"
                + Environment.NewLine + string.Join(Environment.NewLine, violaciones));
        }
    }

    [Fact]
    public void T4_ElMapaCongeladoPorEmpresa_NoSeConsumeEnProduccion()
    {
        var raiz = Insumos.Raiz();
        var produccion = new[]
        {
            Path.Combine(raiz, "Remuneracion.Core"),
            Path.Combine(raiz, "Remuneracion.Infrastructure"),
            Path.Combine(raiz, "Remuneracion.WinForms"),
            Path.Combine(raiz, "Remuneracion.Cli")
        };

        // Solo las tablas EDITABLES por (empresa, ASE) son el mecanismo retirado (H1.2/D7). La
        // inventario de fórmulas protegidas que vive en la misma clase NO es resolución por empresa.
        var referencias = produccion
            .SelectMany(dir => Directory.EnumerateFiles(dir, "*.cs", SearchOption.AllDirectories))
            .Where(archivo => !archivo.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.Ordinal))
            .Where(archivo => !archivo.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.Ordinal))
            .Where(archivo => !string.Equals(Path.GetFileName(archivo), "WorkbookLeafCellMapPorEmpresa.cs", StringComparison.OrdinalIgnoreCase))
            .Select(archivo => (Archivo: archivo, Codigo: string.Join('\n', SinComentarios(File.ReadAllText(archivo)))))
            .Where(par => par.Codigo.Contains("EditablesR1PorEmpresa", StringComparison.Ordinal)
                || par.Codigo.Contains("EditablesR2PorEmpresa", StringComparison.Ordinal)
                || par.Codigo.Contains("EditablesR4PorEmpresa", StringComparison.Ordinal))
            .Select(par => Path.GetRelativePath(raiz, par.Archivo))
            .ToList();

        // Plan 55 (H1.2, D7/CA7): las tablas editables quedan LEGACY (solo tests históricos).
        Assert.Empty(referencias);
    }

    [Fact]
    public void T4_Injerto_ListaPorAse_Enrojece()
    {
        const string injerto = """
            var destino = aseId switch
            {
                1 => "F33",
                4 => "D110",
                _ => "F0"
            };
            """;

        Assert.NotEmpty(Analizar(injerto));
    }

    [Fact]
    public void T4_Injerto_LiteralDeEmpresa_Enrojece()
    {
        Assert.NotEmpty(Analizar("var etiqueta = \"OCCIDENTE\";"));
        Assert.NotEmpty(Analizar("if (fila.C == \"NUEVO ESQUEMA\") { return; }"));
    }

    [Fact]
    public void T4_Injerto_MapaCongelado_Enrojece()
    {
        Assert.NotEmpty(Analizar("var mapa = WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa;"));
    }

    [Fact]
    public void T4_ElAnalizador_NoEnrojeceConLaMaquinariaReusada()
    {
        // Allow-list: la maquinaria citada por la doctrina (sinonimia, secciones, columnas por
        // encabezado, catálogo) NO es una violación.
        const string permitido = """
            var empresa = EmpresaFacturacion.Catalogo.FirstOrDefault(e => SinonimosEmpresaR1.SonMismaEmpresa(e.LabelTemplate, etiqueta));
            var seccion = ContratoSeccionPeriodFreeR1.SeccionPorCierre(filas, i);
            var canal = ContratoSeccionPeriodFreeR1.CanalEfectivo(filas, i);
            var ase = AseFactory.DesdeId(aseId);
            """;

        Assert.Empty(Analizar(permitido));
    }

    // ── analizador (helper del candado) ────────────────────────────────────────────────────────

    /// <summary>
    /// Analiza el código (sin comentarios) y devuelve las violaciones del candado: ramas/listas por
    /// ASE, literales de empresa y consumo del mapa congelado por (empresa, ASE).
    /// </summary>
    internal static IReadOnlyList<string> Analizar(string fuente)
    {
        var violaciones = new List<string>();
        var lineas = SinComentarios(fuente);

        var ramaPorAse = new[]
        {
            new Regex(@"\baseId\s*switch\b", RegexOptions.IgnoreCase),
            new Regex(@"\bAse\.Id\s*(==|switch|is)\b", RegexOptions.IgnoreCase),
            new Regex(@"\bcase\s+[1-5]\s*:", RegexOptions.IgnoreCase),
            new Regex(@"\[\s*\(?\s*[1-5]\s*,\s*[1-5]\s*\)?\s*\]\s*=", RegexOptions.IgnoreCase)
        };

        for (var i = 0; i < lineas.Count; i++)
        {
            var linea = lineas[i];
            foreach (var patron in ramaPorAse)
            {
                if (patron.IsMatch(linea))
                {
                    violaciones.Add($"línea {i + 1}: lista/rama por ASE → {linea.Trim()}");
                }
            }

            if (linea.Contains("WorkbookLeafCellMapPorEmpresa", StringComparison.Ordinal))
            {
                violaciones.Add($"línea {i + 1}: consume el mapa congelado por (empresa, ASE) → {linea.Trim()}");
            }

            foreach (Match literal in Regex.Matches(linea, "\"([^\"]*)\""))
            {
                foreach (var rotulo in RotulosDeEmpresa)
                {
                    if (SinonimosEmpresaR1.SonMismaEmpresa(literal.Groups[1].Value, rotulo))
                    {
                        violaciones.Add($"línea {i + 1}: literal de empresa '{literal.Groups[1].Value}' → {linea.Trim()}");
                    }
                }
            }
        }

        return violaciones;
    }

    /// <summary>Quita comentarios de línea y de bloque (el candado juzga CÓDIGO, no prosa).</summary>
    private static List<string> SinComentarios(string fuente)
    {
        var sinBloque = Regex.Replace(fuente, @"/\*.*?\*/", string.Empty, RegexOptions.Singleline);
        return sinBloque
            .Split('\n')
            .Select(linea =>
            {
                var indice = linea.IndexOf("//", StringComparison.Ordinal);
                return indice >= 0 ? linea[..indice] : linea;
            })
            .ToList();
    }

    private static string Region(string ruta, string desde, string hasta)
    {
        var texto = File.ReadAllText(ruta);
        var inicio = texto.IndexOf(desde, StringComparison.Ordinal);
        Assert.True(inicio >= 0, $"No se encontró la región '{desde}' en {ruta}.");
        var fin = texto.IndexOf(hasta, inicio, StringComparison.Ordinal);
        Assert.True(fin > inicio, $"No se encontró el cierre '{hasta}' en {ruta}.");

        return texto[inicio..fin];
    }

    private static string Ruta(string relativo) => Path.Combine(Insumos.Raiz(), relativo);
}
