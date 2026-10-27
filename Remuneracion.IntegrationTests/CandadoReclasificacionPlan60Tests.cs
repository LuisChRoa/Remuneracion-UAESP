using System.Text.RegularExpressions;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 60 (H2.4/H3.4 — §6 G7, CA-R6): CANDADO ANTI-QUEMADO de la reclasificación
/// <c>Deb/Cred</c> → <c>Deb/Cred-otros</c>.
///
/// <para>La regla es GENERAL de la pestaña <c>Rem. Anticipos R2</c> (punto 6 del texto de la
/// administradora): no puede quedar atada a Occidente, al canal «O», a un ASE ni a un período. El
/// candado juzga el CÓDIGO (los comentarios y la documentación XML se descartan, donde citar el caso
/// como evidencia es legítimo) de los dos puntos donde vive la regla:</para>
/// <list type="number">
/// <item>el núcleo puro <c>Remuneracion.Core/Rules/ReclasificadorDebCredR2.cs</c>;</item>
/// <item>el punto único del writer — la región de <c>EscribirCeldasDetalleR2R4</c> (aplicación + traza).</item>
/// </list>
///
/// <para>Falla si aparece: un rótulo de empresa/ASE del negocio como literal, una comparación contra
/// un literal de canal (una o dos letras), una comparación contra un id numérico de ASE, una
/// coordenada de celda/columna del par (<c>"Q…"</c> / <c>"R…"</c> / <c>== "Q"</c>) o un código de
/// período. Los injertos demuestran que el candado no es decorativo.</para>
/// </summary>
public sealed class CandadoReclasificacionPlan60Tests
{
    /// <summary>Rótulos del negocio que NUNCA pueden gobernar la regla (son datos, no ramas).</summary>
    private static readonly string[] RotulosDeNegocio =
    [
        "OCCIDENTE",
        "ENEL",
        "ENERBIT",
        "NUEVO ESQUEMA",
        "RECIPROCIDAD",
        "EAAB",
        "EAB",
        "LIME",
        "PROMOAMBIENTAL",
        "CIUDAD LIMPIA",
        "BOGOTA LIMPIA",
        "BOGOTÁ LIMPIA",
        "AREA LIMPIA",
        "ÁREA LIMPIA",
    ];

    /// <summary>Comparación contra un literal de canal (una o dos letras/dígitos).</summary>
    private static readonly Regex ComparacionDeCanal = new(
        @"(?:==|!=|Equals\s*\()\s*""[A-Za-z0-9]{1,2}""",
        RegexOptions.Compiled);

    /// <summary>Comparación contra un id numérico de ASE.</summary>
    private static readonly Regex ComparacionDeAse = new(
        @"\b(?:ase|aseId|AseId|numeroAse)\b[^\n]{0,24}(?:==|!=)\s*\d+",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>Literal de celda/columna del par: <c>"Q110"</c>, <c>"R147"</c>, <c>"Q"</c>, <c>"R"</c>.</summary>
    private static readonly Regex LiteralDeCeldaDelPar = new(
        @"""[QR]\d*""",
        RegexOptions.Compiled);

    /// <summary>Código de período (AAAAMM / AAAAMMQ / sello de fecha).</summary>
    private static readonly Regex LiteralDePeriodo = new(@"(?<!\d)20\d{4,6}(?!\d)", RegexOptions.Compiled);

    [Fact]
    public void G7_LaReglaYSuPuntoUnico_NoTienenQuemados()
    {
        var raiz = Insumos.Raiz();
        var objetivo = ObjetivosDeLaRegla(raiz);

        Assert.Equal(2, objetivo.Count);
        Assert.All(objetivo, o => Assert.False(string.IsNullOrWhiteSpace(o.Codigo), $"No se pudo extraer el código de '{o.Archivo}'."));

        var violaciones = objetivo
            .SelectMany(o => Analizar(o.Codigo, o.Archivo))
            .ToList();

        Assert.True(
            violaciones.Count == 0,
            "G-QUEMADO/§1.1: la reclasificación Deb/Cred -> Deb/Cred-otros quedó atada a un dato del caso"
            + " (canal, empresa, ASE, período o coordenada de celda). La regla es GENERAL de 'Rem. Anticipos R2':"
            + Environment.NewLine
            + string.Join(Environment.NewLine, violaciones));
    }

    [Fact]
    public void G7_Injerto_CanalEmpresaAseYPeriodo_Enrojecen()
    {
        Assert.NotEmpty(Analizar("""if (canal == "O") { return; }""", "Regla.cs"));
        Assert.NotEmpty(Analizar("""if (empresa == "OCCIDENTE") { return; }""", "Regla.cs"));
        Assert.NotEmpty(Analizar("""if (ase == 2) { return; }""", "Regla.cs"));
        Assert.NotEmpty(Analizar("""if (aseId != 3) { return; }""", "Regla.cs"));
        Assert.NotEmpty(Analizar("""var ancla = "2026101";""", "Regla.cs"));
    }

    [Fact]
    public void G7_Injerto_CoordenadasDelPar_Enrojecen()
    {
        Assert.NotEmpty(Analizar("""var celda = "Q110";""", "Regla.cs"));
        Assert.NotEmpty(Analizar("""var celda = "R147";""", "Regla.cs"));
        Assert.NotEmpty(Analizar("""var columna = "Q";""", "Regla.cs"));
        Assert.NotEmpty(Analizar("""var columna = "R";""", "Regla.cs"));
    }

    [Fact]
    public void G7_NoEnrojeceConLoLegitimo()
    {
        // Comentarios / XML-doc: citar el caso como evidencia NO es quemar la regla.
        const string documentacion = """
            /// <summary>Ejemplo medido: OCCIDENTE, canal O, ASE 2, período 2026101 (evidencia §0.4).</summary>
            // if (canal == "O") era el atajo prohibido: la etiqueta O se repite en dos cuadros.
            var umbral = UmbralReclasificacionDebCred;
            var puerto = ComponenteDebCred;
            var hoja = HojaRemAnticiposR2;
            """;

        Assert.Empty(Analizar(documentacion, "Regla.cs"));
    }

    // ── analizador (helper del candado) ────────────────────────────────────────────────────────

    /// <summary>
    /// Analiza el CÓDIGO (sin comentarios) y devuelve las violaciones de la clase «dato del caso
    /// quemado en la regla».
    /// </summary>
    internal static IReadOnlyList<string> Analizar(string fuente, string archivo)
    {
        var violaciones = new List<string>();
        var lineas = SinComentarios(fuente);

        for (var i = 0; i < lineas.Count; i++)
        {
            var linea = lineas[i];

            foreach (var rotulo in RotulosDeNegocio)
            {
                if (linea.Contains($"\"{rotulo}\"", StringComparison.OrdinalIgnoreCase))
                {
                    violaciones.Add($"{archivo}:{i + 1}: rótulo de negocio quemado ('{rotulo}') → {linea.Trim()}");
                }
            }

            if (ComparacionDeCanal.IsMatch(linea))
            {
                violaciones.Add($"{archivo}:{i + 1}: comparación contra un literal de canal → {linea.Trim()}");
            }

            if (ComparacionDeAse.IsMatch(linea))
            {
                violaciones.Add($"{archivo}:{i + 1}: comparación contra un id de ASE → {linea.Trim()}");
            }

            if (LiteralDeCeldaDelPar.IsMatch(linea))
            {
                violaciones.Add($"{archivo}:{i + 1}: coordenada de celda/columna del par → {linea.Trim()}");
            }

            if (LiteralDePeriodo.IsMatch(linea))
            {
                violaciones.Add($"{archivo}:{i + 1}: código de período quemado → {linea.Trim()}");
            }
        }

        return violaciones;
    }

    /// <summary>
    /// Los dos objetivos del candado: el núcleo puro completo y la región del punto único del writer
    /// (aplicación de la regla + traza de auditoría). Nunca se juzga el writer entero: su geometría
    /// (letras de columna de los rótulos A–D) es legítima.
    /// </summary>
    private static IReadOnlyList<(string Archivo, string Codigo)> ObjetivosDeLaRegla(string raiz)
    {
        var regla = Path.Combine(raiz, "Remuneracion.Core", "Rules", "ReclasificadorDebCredR2.cs");
        Assert.True(File.Exists(regla), $"Falta el núcleo de la regla: {regla}");

        var writer = Path.Combine(raiz, "Remuneracion.Infrastructure", "Excel", "OpenXmlPlantillaWriter.cs");
        Assert.True(File.Exists(writer), $"Falta el writer: {writer}");

        const string inicio = "private static void EscribirCeldasDetalleR2R4(";
        const string fin = "private static void EscribirDetalleR4(";
        var texto = File.ReadAllText(writer);
        var desde = texto.IndexOf(inicio, StringComparison.Ordinal);
        var hasta = texto.IndexOf(fin, StringComparison.Ordinal);
        Assert.True(desde >= 0 && hasta > desde, "No se pudo delimitar la región del punto único de la reclasificación en el writer.");

        return
        [
            ("Remuneracion.Core/Rules/ReclasificadorDebCredR2.cs", File.ReadAllText(regla)),
            ("Remuneracion.Infrastructure/Excel/OpenXmlPlantillaWriter.cs [EscribirCeldasDetalleR2R4..EscribirDetalleR4]", texto[desde..hasta]),
        ];
    }

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
}
