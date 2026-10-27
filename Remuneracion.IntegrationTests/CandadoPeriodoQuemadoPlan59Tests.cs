using System.Text.RegularExpressions;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (H5, §0.1/§0.2 + §2-D5): CANDADO ANTI-QUEMADO DE PERÍODO.
///
/// Mandato del Ingeniero (textual): «¿La app tiene reglas o cosas quemadas por período? […] Eso es
/// inadmisible, eso no lo permitimos». El período es un DATO del dominio
/// (<c>Periodo.CodigoCompleto</c>) y el nombre de las hojas que lo llevan se LEE del workbook
/// (<c>HojasPeriodo</c>); nunca se escribe en el código.
///
/// Falla si en los proyectos de PRODUCCIÓN (Core, Infrastructure, WinForms, Cli) un literal de cadena
/// contiene:
/// <list type="number">
/// <item>un código de período (<c>20</c> + 4..6 dígitos: AAAAMM, AAAAMMQ o el sello de fecha);</item>
/// <item>un nombre de hoja con sufijo de período (<c>DetRetri2026…</c>, <c>DetValiRetri2026…</c>,
/// <c>Informe AFaseo Recaudo 2026…</c>).</item>
/// </list>
/// Las excepciones son una LISTA CERRADA con motivo declarado (no un liberalismo): cada entrada dice
/// archivo, fragmento y por qué no es una regla por período. La prueba de injerto exige que un literal
/// inyectado a propósito ENROJEZCA, para que el candado no pueda volverse decorativo.
///
/// El análisis juzga CÓDIGO: los comentarios (incluida la documentación XML, donde se citan períodos
/// reales como evidencia) se descartan antes de evaluar.
/// </summary>
public sealed class CandadoPeriodoQuemadoPlan59Tests
{
    /// <summary>Raíces de producción (los tests no juzgan su propio código por este candado).</summary>
    private static readonly string[] ProyectosDeProduccion =
    [
        "Remuneracion.Core",
        "Remuneracion.Infrastructure",
        "Remuneracion.WinForms",
        "Remuneracion.Cli"
    ];

    /// <summary>
    /// Código de período en un literal: <c>20</c> + 4..6 dígitos (AAAAMM / AAAAMMQ / sello de fecha),
    /// delimitado para no confundirlo con importes, ids o números de plan.
    /// </summary>
    private static readonly Regex LiteralDePeriodo = new(@"(?<!\d)20\d{4,6}(?!\d)", RegexOptions.Compiled);

    /// <summary>Nombre de hoja con sufijo de período.</summary>
    private static readonly Regex NombreDeHojaConPeriodo = new(
        @"(?:DetRetri|DetValiRetri)\s*\d{4,7}|Informe\s+AFaseo\s+Recaudo\s+20\d\d",
        RegexOptions.Compiled);

    /// <summary>
    /// Lista CERRADA de excepciones declaradas: (archivo relativo, fragmento del literal, motivo).
    /// Un literal solo queda exento si el archivo coincide Y el literal contiene el fragmento.
    /// </summary>
    private static readonly (string Archivo, string Fragmento, string Motivo)[] Exentos =
    [
        (
            Path.Combine("Remuneracion.Cli", "OpcionesCli.cs"),
            "p. ej. 2026071",
            "Texto de AYUDA del CLI: documenta el FORMATO del argumento (AAAAMM|AAAAMMQ) con un ejemplo; no decide comportamiento ni selecciona hoja, celda ni rama.")
    ];

    /// <summary>
    /// El nombre del mapa LEGACY de la foto BCE no puede aparecer en producción: si alguien lo vuelve a
    /// consumir, el candado lo dice antes de que la foto se convierta otra vez en fuente de coordenadas.
    /// </summary>
    private const string FotoLegacyBce = "BalanceScGoldenV1";

    [Fact]
    public void H5_Produccion_SinLiteralesDePeriodoNiHojasConSufijoDePeriodo()
    {
        var raiz = Insumos.Raiz();
        var violaciones = new List<string>();

        foreach (var archivo in ArchivosDeProduccion(raiz))
        {
            var relativo = Path.GetRelativePath(raiz, archivo);
            violaciones.AddRange(Analizar(File.ReadAllText(archivo), relativo));
        }

        Assert.True(
            violaciones.Count == 0,
            "G-QUEMADO/§0.2: hay literales de período en producción (el período es un DATO del dominio;"
            + " el sufijo de hoja se lee del workbook con HojasPeriodo):"
            + Environment.NewLine
            + string.Join(Environment.NewLine, violaciones));
    }

    [Fact]
    public void H5_LaFotoLegacyDelBalance_NoSeConsumeEnProduccion()
    {
        var raiz = Insumos.Raiz();
        var consumidores = ArchivosDeProduccion(raiz)
            .Where(archivo => string.Join('\n', SinComentarios(File.ReadAllText(archivo)))
                .Contains(FotoLegacyBce, StringComparison.Ordinal))
            .Select(archivo => Path.GetRelativePath(raiz, archivo))
            .ToList();

        Assert.True(
            consumidores.Count == 0,
            $"La foto congelada '{FotoLegacyBce}' es material histórico de fixtures: la consumen {string.Join(", ", consumidores)}."
            + " El camino productivo verifica el balance por identidad/forma (VerificadorBalanceScDerivado).");
    }

    [Fact]
    public void H5_Injerto_LiteralDePeriodo_Enrojece()
    {
        Assert.NotEmpty(Analizar("""var hoja = "DetRetri" + "2026071";""", "X.cs"));
        Assert.NotEmpty(Analizar("""if (periodo.CodigoCompleto == "2026101") { return; }""", "X.cs"));
        Assert.NotEmpty(Analizar("""var ancla = "2026072";""", "X.cs"));
    }

    [Fact]
    public void H5_Injerto_NombreDeHojaConSufijo_Enrojece()
    {
        Assert.NotEmpty(Analizar("""var hoja = "DetRetri2026071";""", "X.cs"));
        Assert.NotEmpty(Analizar("""var hoja = "DetValiRetri2026101";""", "X.cs"));
        Assert.NotEmpty(Analizar("""var hoja = "Informe AFaseo Recaudo 202607-2";""", "X.cs"));
    }

    [Fact]
    public void H5_ElAnalizador_NoEnrojeceConLoLegitimo()
    {
        // Comentarios y documentación XML: citar un período real como evidencia NO es quemarlo.
        const string documentacion = """
            /// <summary>Medido contra el oráculo: el R10 del período 2026071 declara D14 = Σ filas.</summary>
            // Plan 59: la hoja DetRetri2026072 dejó de estar congelada.
            var codigo = periodo.CodigoCompleto;   // p. ej. 2026101
            var nombre = HojasPeriodo.DetRetri(parte, periodo);
            var mensaje = "--periodo <AAAAMM|AAAAMMQ>";
            """;

        Assert.Empty(Analizar(documentacion, "X.cs"));
    }

    [Fact]
    public void H5_LaExcepcionDeclarada_CubreSoloElTextoDeAyuda()
    {
        const string ayuda =
            """--periodo AAAAMM[Q]  Período (p. ej. 2026071). Si usa AAAAMM, --quincena es obligatoria.""";
        var archivo = Path.Combine("Remuneracion.Cli", "OpcionesCli.cs");

        Assert.Empty(Analizar($"var texto = \"\"\"{ayuda}\"\"\";", archivo));

        // La misma frase en otro archivo NO está exenta: la excepción es por archivo declarado.
        Assert.NotEmpty(Analizar($"var texto = \"{ayuda}\";", Path.Combine("Remuneracion.Core", "Periodo.cs")));
    }

    // ── analizador (helper del candado) ────────────────────────────────────────────────────────

    /// <summary>
    /// Analiza CÓDIGO (sin comentarios) y devuelve las violaciones: ocurrencias de un código de período
    /// y de un nombre de hoja con sufijo de período fuera de la lista de excepciones declaradas.
    ///
    /// <para>Se juzga LÍNEA a línea sobre el texto sin comentarios, y no solo el contenido de las
    /// cadenas entrecomilladas: así el candado también ve los literales CRUDOS de varias líneas
    /// (<c>"""…"""</c>, donde vive el texto de ayuda del CLI), que un escáner por comillas dejaría
    /// pasar. En CÓDIGO —fuera de comentarios— una secuencia <c>20…</c> de 6 a 8 dígitos solo puede
    /// venir de un literal: importes con esos dígitos no existen en el dominio (los montos van en
    /// <c>decimal</c> con separador de miles y los ids de plan son de 3 dígitos).</para>
    /// </summary>
    internal static IReadOnlyList<string> Analizar(string fuente, string archivo)
    {
        var violaciones = new List<string>();
        var lineas = SinComentarios(fuente);

        for (var i = 0; i < lineas.Count; i++)
        {
            var linea = lineas[i];

            if (EsExento(archivo, linea))
            {
                continue;
            }

            if (LiteralDePeriodo.IsMatch(linea))
            {
                violaciones.Add(
                    $"{archivo}:{i + 1}: código de período quemado → {linea.Trim()}");
            }

            if (NombreDeHojaConPeriodo.IsMatch(linea))
            {
                violaciones.Add(
                    $"{archivo}:{i + 1}: nombre de hoja con sufijo de período → {linea.Trim()}");
            }
        }

        return violaciones;
    }

    private static bool EsExento(string archivo, string texto) =>
        Exentos.Any(exento =>
            string.Equals(exento.Archivo, archivo, StringComparison.OrdinalIgnoreCase)
            && texto.Contains(exento.Fragmento, StringComparison.Ordinal));

    /// <summary>Quita comentarios de línea y de bloque: el candado juzga CÓDIGO, no prosa.</summary>
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

    private static IEnumerable<string> ArchivosDeProduccion(string raiz) =>
        ProyectosDeProduccion
            .Select(dir => Path.Combine(raiz, dir))
            .Where(Directory.Exists)
            .SelectMany(dir => Directory.EnumerateFiles(dir, "*.cs", SearchOption.AllDirectories))
            .Where(archivo => !archivo.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.Ordinal))
            .Where(archivo => !archivo.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.Ordinal));
}
