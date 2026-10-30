using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 64 (H4.3/H4.4 — REQ-09 / CA-05 / R-07): GOLDEN de la hoja <c>ANT EXT-REV</c> por bloque/ASE
/// contra los CUATRO manuales del administrativo (Julio-2 <c>2026072</c>, Septiembre-1 <c>2026091</c>,
/// Septiembre-2 <c>2026092</c>, Octubre-1 <c>2026101</c>) y guarda anti-falso-verde del comparador.
///
/// <para><b>Oráculo (no se inventa ningún número).</b> El manual del administrativo es el golden: el
/// destino de cada valor se resuelve con la MISMA autoridad de producción
/// (<see cref="ResolvedorDestinoAntExtRev"/>), nunca con una tabla de letras del fixture (D-05). Se
/// comparan las 100 celdas de valor ±0,5 y las 130 fórmulas de la hoja fórmula-vs-fórmula.</para>
///
/// <para><b>Vacío ≡ 0 (declarado, §0.6 / T0 §5.3 + §8-H1).</b> El único vacío declarado es
/// <c>2026092!J20</c>: el fixture lo fija por PERÍODO y exige que ningún otro manual deje una celda de
/// la rejilla en blanco. Un vacío nuevo (o el declarado en otro período) enrojece — no se generaliza a
/// «vacío = 0» en toda la hoja.</para>
///
/// <para><b>f-vs-f de verdad (M7 del §0.6).</b> Los dos workbooks pueden representar la MISMA fórmula
/// compartida con el maestro en celdas distintas (Excel re-guarda el manual y mueve el grupo
/// <c>t="shared"</c>: medido 14 celdas en Septiembre-1 y Octubre-1, 0 en Julio-2 y Septiembre-2). Por
/// eso se compara el TEXTO EFECTIVO de cada fórmula: las seguidoras se expanden desde su maestro
/// trasladando las referencias relativas. Así la comparación es de contenido, no de representación, y
/// no se abre ninguna tolerancia para fórmulas.</para>
///
/// <para><b>Anti-falso-verde (H4.4 / R-07).</b> Dos guardas: (a) el clasificador de brechas de
/// <see cref="ComparadorSalidaVsManualTests"/> ya NO declara la hoja (solo <c>ANTICIPOS USUARIOS</c>
/// sigue declarada) y el comparador no la excluye por forma; (b) al injertar una divergencia real en
/// una celda de valor de la salida, el comparador DEBE reportarla. Si alguien re-introduce la
/// exclusión, ambas enrojecen.</para>
///
/// Sin emojis. Tolerancia ±0,5. Excel cerrado durante la corrida.
/// </summary>
public sealed class AntExtRevComparadorH4Plan64Tests
{
    private const decimal Tolerancia = 0.5m;

    private static readonly string Hoja = ResolvedorDestinoAntExtRev.Hoja;

    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Prns = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>Fórmula con referencias A1 relativas/absolutas (para trasladar una compartida).</summary>
    private static readonly Regex ReferenciaA1 = new(
        @"(?<![A-Za-z0-9_])(?<ca>\$?)(?<col>[A-Z]{1,3})(?<ra>\$?)(?<fila>\d+)",
        RegexOptions.Compiled);

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Rutas explícitas (prohibido Insumos.Raiz() en lo nuevo)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static string RaizRepositorio()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new InvalidOperationException($"No se encontró AGENTS.md partiendo de '{AppContext.BaseDirectory}'.");
    }

    private static string Ruta(params string[] partes) => Path.Combine([RaizRepositorio(), .. partes]);

    private static string PlantillaV4() =>
        Ruta("Docs", "Plantilla estandarizada", "v4", "Formato Remuneracion 2026xx-x Total v4.xlsx");

    /// <summary>Los cuatro manuales del administrativo que el plan fija como goldens (REQ-09).</summary>
    private static readonly (string Periodo, string[] Ruta)[] Manuales =
    [
        ("2026072", ["Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx"]),
        ("2026091", ["Docs", "Prueba Septiembre 1", "Resultados", "Administrativo", "Remuneracion 202609-1 Total.xlsx"]),
        ("2026092", ["Docs", "Prueba Septiembre-2", "Resultados", "Administrativo", "Remuneracion 202609-2 Total.xlsx"]),
        ("2026101", ["Docs", "Prueba Octubre-1", "Resultados", "Resultado administrativo", "Remuneracion 202610-1 Total.xlsx"])
    ];

    /// <summary>Carpeta de insumos por período (misma estructura de 5 carpetas ASE).</summary>
    private static readonly (string Periodo, string Carpeta)[] CarpetasInsumosPorPeriodo =
    [
        ("2026072", "Prueba Julio-2"),
        ("2026091", "Prueba Septiembre 1"),
        ("2026092", "Prueba Septiembre-2"),
        ("2026101", "Prueba Octubre-1")
    ];

    /// <summary>
    /// ÚNICO vacío declarado de la rejilla en los cuatro manuales (§0.6 «Adicional H1-manual», T0 §5.3
    /// y §8-H1): <c>2026092!J20</c>. Cualquier otro se considera defecto del golden.
    /// </summary>
    private static readonly (string Periodo, string Celda)[] VaciosDeclarados = [("2026092", "J20")];

    private static string Manual(string periodo) => Ruta(Manuales.Single(m => m.Periodo == periodo).Ruta);

    private static string CarpetaInsumos(string periodo) =>
        Ruta("Docs", CarpetasInsumosPorPeriodo.Single(p => p.Periodo == periodo).Carpeta, "Insumos");

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Corrida del período (misma vía que el operador: ProcesadorPeriodo sobre los insumos en disco)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Ejecuta el período multi-ASE con los insumos REALES hacia un directorio temporal fresco (jamás
    /// sobre <c>Docs/</c>). Es la misma composición que el CLI cinco-ASE: así el golden lee lo que el
    /// writer dejó en el workbook (G-WORKBOOK), no el dominio.
    /// </summary>
    private static string EjecutarPeriodo(string periodo)
    {
        var dir = Path.Combine(Path.GetTempPath(), "p64-h4-comparador-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var salida = Path.Combine(dir, Periodo.Parse(periodo).NombreArchivo);

        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());

        procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Periodo.Parse(periodo),
            CarpetaPeriodo = CarpetaInsumos(periodo),
            RutaPlantilla = PlantillaV4(),
            RutaSalida = salida
        });

        Assert.True(File.Exists(salida), "El pipeline debe haber dejado el workbook de salida.");
        return salida;
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H4.3 — Golden por bloque/ASE contra los 4 manuales (±0,5 + f-vs-f)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// CA-05: para cada uno de los cuatro manuales, la salida de la app coincide con el manual en las
    /// 100 celdas de valor resueltas por bloque/ASE/fila/reporte (±0,5) y en las 130 fórmulas de la
    /// hoja (texto efectivo, f-vs-f). El único vacío admitido en el manual es el declarado
    /// <c>2026092!J20</c>, que cuenta como 0.
    /// </summary>
    [Theory]
    [InlineData("2026072")]
    [InlineData("2026091")]
    [InlineData("2026092")]
    [InlineData("2026101")]
    public void AntExtRev_PorBloqueYAse_CoincideConElManual_EnValoresYFormulas(string periodo)
    {
        var manual = Manual(periodo);
        Assert.True(File.Exists(manual), $"Falta el manual del administrativo de {periodo}: {manual}");

        var salida = EjecutarPeriodo(periodo);

        // Autoridad del destino: la producción resuelve la plantilla que escribió y el manual por separado;
        // ambas rejillas deben ser la MISMA (si el manual movió una etiqueta, esto lo delata antes del valor).
        var destinos = DestinosDe(salida);
        Assert.Equal(100, destinos.Count);
        Assert.Equal(
            destinos.Select(d => d.Celda).OrderBy(c => c, StringComparer.Ordinal),
            DestinosDe(manual).Select(d => d.Celda).OrderBy(c => c, StringComparer.Ordinal));

        var valoresSalida = LeerValoresDeHoja(salida, Hoja);
        var valoresManual = LeerValoresDeHoja(manual, Hoja);

        // (a) La salida SIEMPRE trae las 100 celdas como valor numérico (H2/REQ-07: values-only, M1).
        foreach (var destino in destinos)
        {
            Assert.True(
                valoresSalida.ContainsKey(destino.Celda),
                $"La salida de {periodo} no escribió {destino.Celda} ({Identidad(destino)}).");
        }

        // (b) Vacío ≡ 0 SOLO en el vacío declarado: ningún otro manual deja su rejilla en blanco.
        var ausentesEnElManual = destinos
            .Where(destino => !valoresManual.ContainsKey(destino.Celda))
            .Select(destino => destino.Celda)
            .OrderBy(celda => celda, StringComparer.Ordinal)
            .ToArray();
        var vaciosEsperados = VaciosDeclarados
            .Where(vacio => vacio.Periodo == periodo)
            .Select(vacio => vacio.Celda)
            .OrderBy(celda => celda, StringComparer.Ordinal)
            .ToArray();

        Assert.True(
            vaciosEsperados.SequenceEqual(ausentesEnElManual, StringComparer.OrdinalIgnoreCase),
            $"Los vacíos del manual de {periodo} no son los declarados (§0.6: solo 2026092!J20). " +
            $"Esperado=[{string.Join(", ", vaciosEsperados)}] observado=[{string.Join(", ", ausentesEnElManual)}].");

        // Anti-vacuidad (S6): la hoja no puede pasar comparando "0 contra 0" — el insumo trae dinero real
        // (T0 §1.4: montos por bloque/ASE distintos de cero en los cuatro períodos).
        Assert.Contains(destinos, destino => ValorDeCelda(valoresSalida, destino.Celda) != 0m);
        Assert.Contains(destinos, destino => ValorDeCelda(valoresManual, destino.Celda) != 0m);

        // (c) ±0,5 por celda contra el manual (ausente del manual = 0 solo con la guarda de (b)).
        var divergencias = new List<string>();
        foreach (var destino in destinos)
        {
            var valorApp = ValorDeCelda(valoresSalida, destino.Celda);
            var valorManual = ValorDeCelda(valoresManual, destino.Celda);
            if (Math.Abs(valorApp - valorManual) > Tolerancia)
            {
                divergencias.Add(
                    $"{destino.Celda} ({Identidad(destino)}) app={valorApp.ToString(CultureInfo.InvariantCulture)} " +
                    $"manual={valorManual.ToString(CultureInfo.InvariantCulture)} " +
                    $"delta={(valorApp - valorManual).ToString(CultureInfo.InvariantCulture)}");
            }
        }

        Assert.True(
            divergencias.Count == 0,
            $"ANT EXT-REV de {periodo} diverge del manual en {divergencias.Count} celda(s) (±{Tolerancia}):\n"
            + string.Join("\n", divergencias));

        // (d) f-vs-f: las 130 fórmulas de la hoja por TEXTO EFECTIVO (seguidoras compartidas expandidas).
        var formulasSalida = LeerFormulasEfectivas(salida, Hoja);
        var formulasManual = LeerFormulasEfectivas(manual, Hoja);
        Assert.Equal(130, formulasSalida.Count);
        Assert.Equal(130, formulasManual.Count);

        // Values-only: ninguna de las 100 celdas de valor puede traer fórmula (M1).
        foreach (var destino in destinos)
        {
            Assert.False(
                formulasSalida.ContainsKey(destino.Celda),
                $"La celda de valor {destino.Celda} ({Identidad(destino)}) quedó con fórmula en la salida de {periodo}.");
        }

        var formulasDivergentes = formulasSalida.Keys
            .Union(formulasManual.Keys, StringComparer.OrdinalIgnoreCase)
            .Where(celda => !string.Equals(
                formulasSalida.GetValueOrDefault(celda), formulasManual.GetValueOrDefault(celda), StringComparison.Ordinal))
            .Select(celda => $"{celda}: app=[{formulasSalida.GetValueOrDefault(celda)}] manual=[{formulasManual.GetValueOrDefault(celda)}]")
            .OrderBy(texto => texto, StringComparer.Ordinal)
            .ToArray();

        Assert.True(
            formulasDivergentes.Length == 0,
            $"El texto EFECTIVO de las fórmulas de ANT EXT-REV de {periodo} difiere del manual en " +
            $"{formulasDivergentes.Length} celda(s):\n" + string.Join("\n", formulasDivergentes));
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H4.4 — Anti-falso-verde (R-07): la hoja ya no se excluye ni se declara en bloque
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// R-07 (injerto en rojo): una divergencia REAL en una celda de valor de <c>ANT EXT-REV</c> debe
    /// aparecer en el informe del comparador. Se altera el workbook GENERADO (nunca el manual ni un
    /// monto del insumo) y se exige la divergencia con celda + motivo. Si alguien vuelve a excluir la
    /// hoja —por forma o por hoja entera— este test enrojece: falso verde sobre dinero escrito.
    /// </summary>
    [Fact]
    public void Comparador_ReportaLaDivergenciaAntExtRev_SinExclusion()
    {
        const string periodo = "2026091";
        var manual = Manual(periodo);
        var salida = EjecutarPeriodo(periodo);

        var destinos = DestinosDe(salida);
        var alterada = destinos.Single(destino =>
            destino.AseId == 1
            && string.Equals(destino.EtiquetaBloque, "BANCO DE OCCIDENTE", StringComparison.Ordinal)
            && string.Equals(destino.Fila, ResolvedorDestinoAntExtRev.FilaAplicado, StringComparison.Ordinal)
            && string.Equals(destino.Reporte, ResolvedorDestinoAntExtRev.Reporte3, StringComparison.Ordinal));

        using (var documento = SpreadsheetDocument.Open(salida, true))
        {
            var parte = documento.WorkbookPart!;
            var hoja = HojaDe(parte, Hoja);
            var celda = hoja.Descendants<Cell>()
                .Single(c => string.Equals(c.CellReference?.Value, alterada.Celda, StringComparison.OrdinalIgnoreCase));
            celda.CellValue = new CellValue(
                (LeerNumero(celda) + 1000m).ToString(CultureInfo.InvariantCulture));
            hoja.Save();
        }

        var divergencias = new ComparadorSalidaVsManual(1).Comparar(salida, manual);
        var deLaHoja = divergencias
            .Where(d => ComparadorSalidaVsManual.NormalizarHoja(d.Hoja) == Hoja)
            .ToArray();

        var reportada = deLaHoja.SingleOrDefault(d =>
            string.Equals(d.Celda, alterada.Celda, StringComparison.OrdinalIgnoreCase));
        Assert.True(
            reportada is not null,
            $"El comparador NO reportó la divergencia inyectada en ANT EXT-REV!{alterada.Celda} " +
            $"({Identidad(alterada)}): la hoja se está tapando (R-07). Divergencias de la hoja: " +
            $"[{string.Join(", ", deLaHoja.Select(d => $"{d.Celda}({d.Motivo})"))}].");
        Assert.Equal("literal numérico fuera de tolerancia (±0.5)", reportada!.Motivo);
    }

    /// <summary>
    /// R-07 (estructural): el clasificador de brechas del comparador declara SOLO
    /// <c>ANTICIPOS USUARIOS</c> (100 % fórmulas, no se escribe). Cualquier celda de
    /// <c>ANT EXT-REV</c> —valor, total o banda superior— debe quedar como INESPERADA, y la exclusión
    /// por forma del comparador no puede tocar la hoja.
    /// </summary>
    [Fact]
    public void BrechaDeclarada_DeclaraAnticiposUsuarios_PeroNoAntExtRev()
    {
        var periodo = new Periodo { CodigoAAAAMM = "202609", NumeroQuincena = 1 };

        // Valor de la rejilla, vacío declarado, TOTAL del bloque y banda superior: ningún caso se declara.
        foreach (var celda in new[] { "C11", "J20", "M3", "N4", "C45", "L129" })
        {
            var divergencia = new ComparadorSalidaVsManual.Divergencia(
                Hoja, celda, "42680", "43680", "literal numérico fuera de tolerancia (±0.5)");
            Assert.Null(ComparadorSalidaVsManualTests.BrechaDeclarada(periodo, divergencia));
        }

        // ANTICIPOS USUARIOS sigue siendo non-goal declarado: la app NO la escribe.
        var anticipos = new ComparadorSalidaVsManual.Divergencia(
            "ANTICIPOS USUARIOS", "C12", "=f('ANT EXT-REV'!C4)", "=f('ANT EXT-REV'!C5)", "texto de fórmula distinto");
        var declarada = ComparadorSalidaVsManualTests.BrechaDeclarada(periodo, anticipos);
        Assert.NotNull(declarada);
        Assert.Contains("non-goal", declarada!, StringComparison.Ordinal);

        // Exclusiones POR FORMA: la hoja no está en la lista versionada del comparador.
        var comparador = new ComparadorSalidaVsManual(1);
        foreach (var celda in new[] { "C11", "J20", "M29", "N3", "D45" })
        {
            Assert.False(
                comparador.EsExclusionVersionada(Hoja, celda),
                $"La celda {Hoja}!{celda} está excluida por forma: falso verde sobre dinero escrito (R-07).");
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Utilidades
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static string Identidad(ResolvedorDestinoAntExtRev.Destino destino) =>
        $"{destino.EtiquetaBloque} · ASE {destino.AseId} · {destino.Fila} · {destino.Reporte}";

    private static IReadOnlyList<ResolvedorDestinoAntExtRev.Destino> DestinosDe(string ruta)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        return ResolvedorDestinoAntExtRev.Resolver(documento.WorkbookPart!);
    }

    private static Worksheet HojaDe(WorkbookPart parte, string nombre)
    {
        var workbook = parte.Workbook ?? throw new InvalidOperationException("El workbook no tiene metadata válida.");
        var hoja = workbook.Descendants<Sheet>()
            .Single(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)parte.GetPartById(hoja.Id!.Value!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet válido.");
    }

    private static decimal LeerNumero(Cell celda) =>
        celda.CellValue is null || string.IsNullOrWhiteSpace(celda.CellValue.InnerText)
            ? 0m
            : decimal.Parse(celda.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture);

    private static decimal ValorDeCelda(IReadOnlyDictionary<string, string> celdas, string referencia)
    {
        if (!celdas.TryGetValue(referencia, out var valor))
        {
            return 0m;
        }

        return decimal.TryParse(valor, NumberStyles.Any, CultureInfo.InvariantCulture, out var numero) ? numero : 0m;
    }

    private static XDocument LeerHojaXml(ZipArchive zip, XDocument relaciones, string nombreHoja)
    {
        var workbook = XDocument.Load(zip.GetEntry("xl/workbook.xml")!.Open());
        var hoja = workbook.Root!.Element(Ns + "sheets")!.Elements(Ns + "sheet")
            .SingleOrDefault(s => (string?)s.Attribute("name") == nombreHoja)
            ?? throw new InvalidOperationException($"El workbook no trae la hoja '{nombreHoja}'.");

        var destinos = relaciones.Root!.Elements(Prns + "Relationship")
            .ToDictionary(r => r.Attribute("Id")!.Value, r => r.Attribute("Target")!.Value);
        var destino = destinos[hoja.Attribute(Rns + "id")!.Value];
        var ruta = destino.StartsWith("/", StringComparison.Ordinal)
            ? destino.TrimStart('/')
            : destino.Replace("worksheets/", "xl/worksheets/", StringComparison.Ordinal);
        if (!ruta.StartsWith("xl/", StringComparison.Ordinal))
        {
            ruta = "xl/" + ruta;
        }

        return XDocument.Load(zip.GetEntry(ruta)!.Open());
    }

    /// <summary>Valores literales (&lt;v&gt; con contenido) de una hoja, por referencia A1.</summary>
    private static Dictionary<string, string> LeerValoresDeHoja(string ruta, string nombreHoja)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var relaciones = XDocument.Load(zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open());
        var hoja = LeerHojaXml(zip, relaciones, nombreHoja);

        return hoja.Descendants(Ns + "c")
            .Where(celda => celda.Attribute("r") is not null
                && !string.IsNullOrEmpty(celda.Element(Ns + "v")?.Value))
            .ToDictionary(
                celda => (string)celda.Attribute("r")!,
                celda => celda.Element(Ns + "v")!.Value!,
                StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Fórmulas EFECTIVAS de una hoja (texto sin espacios): las seguidoras de un grupo compartido se
    /// expanden desde su maestro trasladando las referencias relativas (M7 del §0.6). Comparar el texto
    /// efectivo es f-vs-f de contenido: la posición del maestro dentro del grupo compartido no es una
    /// diferencia de fórmula.
    /// </summary>
    private static Dictionary<string, string> LeerFormulasEfectivas(string ruta, string nombreHoja)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var relaciones = XDocument.Load(zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open());
        var hoja = LeerHojaXml(zip, relaciones, nombreHoja);

        var celdas = hoja.Descendants(Ns + "c")
            .Where(celda => celda.Attribute("r") is not null && celda.Element(Ns + "f") is not null)
            .ToArray();

        var maestros = new Dictionary<string, (string Celda, string Texto)>(StringComparer.Ordinal);
        foreach (var celda in celdas)
        {
            var formula = celda.Element(Ns + "f")!;
            var referencia = (string?)formula.Attribute("ref");
            if (!string.IsNullOrEmpty(referencia) && !string.IsNullOrEmpty(formula.Value)
                && formula.Attribute("si")?.Value is { } indice)
            {
                maestros[indice] = ((string)celda.Attribute("r")!, formula.Value);
            }
        }

        var resultado = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var celda in celdas)
        {
            var formula = celda.Element(Ns + "f")!;
            var referenciaCelda = (string)celda.Attribute("r")!;
            var texto = formula.Value;
            if (string.IsNullOrEmpty(texto)
                && string.Equals(formula.Attribute("t")?.Value, "shared", StringComparison.Ordinal)
                && formula.Attribute("si")?.Value is { } indiceSeguidora)
            {
                var maestro = maestros[indiceSeguidora];
                texto = Trasladar(maestro.Texto, maestro.Celda, referenciaCelda);
            }

            resultado[referenciaCelda] = string.Concat(texto.Where(caracter => !char.IsWhiteSpace(caracter)));
        }

        return resultado;
    }

    /// <summary>Traslada las referencias relativas de <paramref name="formula"/> del origen al destino.</summary>
    private static string Trasladar(string formula, string desde, string hasta)
    {
        var (columnaDesde, filaDesde) = Coordenadas(desde);
        var (columnaHasta, filaHasta) = Coordenadas(hasta);

        return ReferenciaA1.Replace(formula, coincidencia =>
        {
            var columna = coincidencia.Groups["col"].Value;
            var fila = int.Parse(coincidencia.Groups["fila"].Value, CultureInfo.InvariantCulture);
            var columnaFinal = coincidencia.Groups["ca"].Value.Length > 0
                ? columna
                : LetraDe(ColumnaDe(columna) + (columnaHasta - columnaDesde));
            var filaFinal = coincidencia.Groups["ra"].Value.Length > 0
                ? fila
                : fila + (filaHasta - filaDesde);

            return $"{coincidencia.Groups["ca"].Value}{columnaFinal}{coincidencia.Groups["ra"].Value}"
                + filaFinal.ToString(CultureInfo.InvariantCulture);
        });
    }

    private static (int Columna, int Fila) Coordenadas(string celda)
    {
        var indice = 0;
        var columna = 0;
        while (indice < celda.Length && char.IsLetter(celda[indice]))
        {
            columna = (columna * 26) + (char.ToUpperInvariant(celda[indice]) - 'A' + 1);
            indice++;
        }

        return (columna, int.Parse(celda[indice..], CultureInfo.InvariantCulture));
    }

    private static int ColumnaDe(string columna)
    {
        var indice = 0;
        foreach (var caracter in columna)
        {
            indice = (indice * 26) + (char.ToUpperInvariant(caracter) - 'A' + 1);
        }

        return indice;
    }

    private static string LetraDe(int columna)
    {
        var letras = string.Empty;
        while (columna > 0)
        {
            var resto = (columna - 1) % 26;
            letras = (char)('A' + resto) + letras;
            columna = (columna - 1) / 26;
        }

        return letras;
    }
}
