using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Cli;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Serilog;
using Serilog.Core;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 64 (H2.2/H2.3, REQ-07 — CA-03): escritura values-only de <c>ANT EXT-REV</c>, lectura de vuelta y
/// corrida e2e por el CLI real.
///
/// <para><b>Oráculos (no se inventa ningún número).</b>
/// <list type="bullet">
/// <item>El manual del administrativo del período es el golden por bloque/ASE; la celda destino se
/// resuelve con la MISMA autoridad de producción (<see cref="ResolvedorDestinoAntExtRev"/>), nunca con
/// una tabla de letras del fixture.</item>
/// <item>Las fórmulas protegidas (bloque superior y TOTAL por columna) se comparan por texto-<c>&lt;f&gt;</c>
/// entre la plantilla v4 y la salida: el diff debe ser vacío (G-FORMULA).</item>
/// <item><c>ANTICIPOS USUARIOS</c> y <c>CONSOLIDADO_TOTAL RECAUDO!L125:L131</c> (M6) no se escriben: se
/// exige que sus fórmulas y sus valores sigan iguales a la plantilla, con <c>fullCalcOnLoad</c> vigente
/// (el recálculo es de Excel, nunca un <c>&lt;v&gt;</c> stale tomado como oráculo).</item>
/// <item>Los casos negativos se fuerzan sobre el workbook GENERADO (una celda de valor alterada) para
/// probar que la lectura de vuelta falla; no se manipula ningún monto real del insumo.</item>
/// </list></para>
/// </summary>
[Collection("Cli")]
public sealed class AntExtRevEscrituraH2Plan64Tests : IDisposable
{
    private const decimal Tolerancia = 0.5m;

    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Prns = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>Plan 61 (D10): la lectura de vuelta conserva el workbook; aquí se limpia el log global del CLI.</summary>
    public void Dispose()
    {
        Log.CloseAndFlush();
        Log.Logger = Logger.None;
    }

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

    private static readonly (string Periodo, string Carpeta)[] Periodos =
    [
        ("2026091", "Prueba Septiembre 1"),
        ("2026101", "Prueba Octubre-1")
    ];

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia"
    ];

    private static string CarpetaInsumos(string periodo) =>
        Ruta("Docs", Periodos.Single(p => p.Periodo == periodo).Carpeta, "Insumos");

    private static readonly (string Periodo, string[] Ruta)[] Manuales =
    [
        ("2026091", ["Docs", "Prueba Septiembre 1", "Resultados", "Administrativo", "Remuneracion 202609-1 Total.xlsx"]),
        ("2026101", ["Docs", "Prueba Octubre-1", "Resultados", "Resultado administrativo", "Remuneracion 202610-1 Total.xlsx"])
    ];

    private static string Manual(string periodo) => Ruta(Manuales.Single(m => m.Periodo == periodo).Ruta);

    private static string Fuente(string carpetaAse, string prefijo, int aseId)
    {
        var archivo = Directory.EnumerateFiles(carpetaAse, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f)
                .StartsWith(prefijo, StringComparison.OrdinalIgnoreCase))
            ?? throw new FileNotFoundException($"Falta '{prefijo}' del ASE {aseId} en {carpetaAse}.");
        return archivo;
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Armado del período con los insumos realmente en disco (misma vía que el orquestador)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static (List<WorkbookLeafInputs> Leafs, ResultadoRemuneracion Resultado) ConstruirPeriodo(string periodo)
    {
        var dominio = Periodo.Parse(periodo);
        var carpetaPeriodo = CarpetaInsumos(periodo);
        var leafReader = new ExcelDataReaderWorkbookLeafInputReader();
        var lector = new ExcelDataReaderRecaudoReader();
        var leafs = new List<WorkbookLeafInputs>();
        var datos = new List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4)>();

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var carpetaAse = Path.Combine(carpetaPeriodo, CarpetasAse[aseId - 1]);
            Assert.True(Directory.Exists(carpetaAse), $"Falta la carpeta del ASE{aseId} en {periodo}: {carpetaAse}");

            var ase = AseFactory.DesdeId(aseId);
            var rutaR1 = Fuente(carpetaAse, "Recaudoporcomponente", aseId);
            var rutaR2 = Fuente(carpetaAse, "RerpoteDetalleSaldosaFavor", aseId);
            var rutaR4 = Fuente(carpetaAse, "Reversi", aseId);

            var leaf = leafReader.LeerLeafInputs(ase, dominio, rutaR1, rutaR2, rutaR4);
            leaf.BalanceSc = leafReader.LeerBalanceSc(
                ase, dominio, Fuente(carpetaAse, "R4-BalanceSubsidioyContribuciones", aseId));

            // H2.1: la lectura del insumo viaja en el leaf y el writer la consume en su pasada atómica.
            leaf.AntExtRev = leafReader.LeerAntExtRev(
                ase, dominio, Fuente(carpetaAse, "RecaudosReversados", aseId));

            leafs.Add(leaf);
            datos.Add((ase, lector.LeerR1(rutaR1), lector.LeerR2(rutaR2), lector.LeerR4(rutaR4)));
        }

        return (leafs, new CalculoRemuneracion().CalcularConsolidado(dominio, datos));
    }

    private static string NuevoDirectorio()
    {
        var dir = Path.Combine(Path.GetTempPath(), "p64-h2-escritura-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static string Generar(string periodo, IReadOnlyList<WorkbookLeafInputs> leafs, ResultadoRemuneracion resultado)
    {
        var dir = NuevoDirectorio();
        var salida = Path.Combine(dir, Periodo.Parse(periodo).NombreArchivo);
        new OpenXmlPlantillaWriter().GenerarWorkbook(PlantillaV4(), salida, resultado, leafs);
        Assert.True(File.Exists(salida), "El writer debe haber dejado el workbook de salida.");
        return salida;
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H2.1/H2.2 — valores escritos, lectura de vuelta y fórmulas protegidas
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// REQ-07/CA-03 (§2.6): por cada fuente PRESENTE el writer pega sus cuatro montos en el bloque de su
    /// etiqueta y en la columna de su ASE; las fuentes ausentes (y los bloques sin fila de catálogo —
    /// HUE-1) conservan el cero del template. La salida completa coincide con el manual del
    /// administrativo por bloque/ASE ±0,5, la lectura de vuelta del propio writer cierra y el diff de
    /// fórmulas de la hoja es VACÍO (bloque superior + TOTAL por columna intactos).
    /// </summary>
    [Theory]
    [InlineData("2026091")]
    [InlineData("2026101")]
    public void Escritura_AntExtRev_IgualAlManualPorBloqueYAse_ConReadBackYFormulasIntactas(string periodo)
    {
        var (leafs, resultado) = ConstruirPeriodo(periodo);
        var salida = Generar(periodo, leafs, resultado);

        // Autoridad del destino: la propia plantilla/manual, resuelta por etiquetas.
        var destinos = DestinosDe(Manual(periodo));
        var celdasManual = LeerValoresDeHoja(Manual(periodo), ResolvedorDestinoAntExtRev.Hoja);
        var celdasSalida = LeerValoresDeHoja(salida, ResolvedorDestinoAntExtRev.Hoja);
        Assert.Equal(100, destinos.Count);

        foreach (var destino in destinos)
        {
            Assert.InRange(
                ValorDeCelda(celdasSalida, destino.Celda) - ValorDeCelda(celdasManual, destino.Celda),
                -Tolerancia,
                Tolerancia);
        }

        // El manual no deja ninguna de las 100 celdas del período en blanco (§0.6: solo 2026092!J20).
        foreach (var destino in destinos)
        {
            Assert.True(
                celdasManual.ContainsKey(destino.Celda),
                $"El manual de {periodo} deja en blanco {destino.Celda} ({destino.EtiquetaBloque} · ASE {destino.AseId} · {destino.Fila} · {destino.Reporte}) y el único vacío declarado es 2026092!J20.");
        }

        // Lectura de vuelta de producción sobre el archivo guardado: 100 celdas ≡ dominio ±0,5.
        using (var documento = SpreadsheetDocument.Open(salida, false))
        {
            OpenXmlPlantillaWriter.VerificarLecturaAntExtRev(documento.WorkbookPart!, leafs);
        }

        // Values-only estricto: ninguna de las 100 celdas de valor es fórmula.
        var formulasSalida = LeerFormulasDeHoja(salida, ResolvedorDestinoAntExtRev.Hoja);
        foreach (var destino in destinos)
        {
            Assert.False(
                formulasSalida.ContainsKey(destino.Celda),
                $"La celda de valor {destino.Celda} quedó con fórmula: la escritura values-only la prohibe.");
        }

        // Fórmulas protegidas byte-idénticas: bloque superior (C3..N3/C4..N4/fila SUM/controles) y TOTAL
        // por columna, comparadas por texto-<f> (incluye si/ref de las compartidas).
        var formulasPlantilla = LeerFormulasDeHoja(PlantillaV4(), ResolvedorDestinoAntExtRev.Hoja);
        AssertFormulasIdenticas(formulasPlantilla, formulasSalida);
        Assert.Equal(130, formulasSalida.Count);
    }

    /// <summary>
    /// H2.1 / D-04: <c>WorkbookLeafInputs.AntExtRev == null</c> es NO-OP por construcción. Se exige el
    /// comportamiento anterior: las 100 celdas de valor quedan en el cero del template y las fórmulas de
    /// la hoja no cambian (nada se escribe, nada se re-resuelve).
    /// </summary>
    [Fact]
    public void NoOp_SinLecturaDelInsumo_DejaLasCienCeldasEnElCeroDelTemplate()
    {
        const string periodo = "2026091";
        var (leafs, resultado) = ConstruirPeriodo(periodo);
        foreach (var leaf in leafs)
        {
            leaf.AntExtRev = null;
        }

        var salida = Generar(periodo, leafs, resultado);

        var destinos = DestinosDe(PlantillaV4());
        var celdasPlantilla = LeerValoresDeHoja(PlantillaV4(), ResolvedorDestinoAntExtRev.Hoja);
        var celdasSalida = LeerValoresDeHoja(salida, ResolvedorDestinoAntExtRev.Hoja);

        foreach (var destino in destinos)
        {
            Assert.Equal(
                ValorDeCelda(celdasPlantilla, destino.Celda),
                ValorDeCelda(celdasSalida, destino.Celda));
        }

        AssertFormulasIdenticas(
            LeerFormulasDeHoja(PlantillaV4(), ResolvedorDestinoAntExtRev.Hoja),
            LeerFormulasDeHoja(salida, ResolvedorDestinoAntExtRev.Hoja));
    }

    /// <summary>
    /// H2.2 / D-08: la lectura de vuelta NO es tautológica. Al alterar una celda de valor del workbook
    /// generado, el verificador de producción falla nombrando hoja + celda + esperado vs escrito con
    /// <c>ERR-ESCRITURA</c> (familia de lectura-de-vuelta: el archivo se conserva como evidencia).
    /// </summary>
    [Fact]
    public void LecturaDeVuelta_Diverge_FallaNombrandoHojaYCelda()
    {
        const string periodo = "2026091";
        var (leafs, resultado) = ConstruirPeriodo(periodo);
        var salida = Generar(periodo, leafs, resultado);

        var destinos = DestinosDe(salida);
        var alterada = destinos.Single(destino =>
            destino.AseId == 1
            && CatalogoFuentesAntExtRev.BuscarPorEtiquetaBloque(destino.EtiquetaBloque) is { } fila
            && string.Equals(fila.EtiquetaBloque, "BANCO DE OCCIDENTE", StringComparison.Ordinal)
            && string.Equals(destino.Fila, "APLICADO", StringComparison.Ordinal)
            && string.Equals(destino.Reporte, "REPORTE 3", StringComparison.Ordinal));

        using (var documento = SpreadsheetDocument.Open(salida, true))
        {
            var parte = documento.WorkbookPart!;
            var hoja = Hoja(parte, ResolvedorDestinoAntExtRev.Hoja);
            var celda = hoja.Descendants<Cell>()
                .Single(c => string.Equals(c.CellReference?.Value, alterada.Celda, StringComparison.OrdinalIgnoreCase));
            celda.CellValue = new CellValue((LeerNumero(celda) + 1000m).ToString(CultureInfo.InvariantCulture));
            hoja.Save();
        }

        var excepcion = Assert.Throws<CalculoInvalidoException>(() =>
        {
            using var documento = SpreadsheetDocument.Open(salida, false);
            OpenXmlPlantillaWriter.VerificarLecturaAntExtRev(documento.WorkbookPart!, leafs);
        });

        Assert.Equal(CodigoError.Escritura, excepcion.Codigo);
        Assert.Contains(ResolvedorDestinoAntExtRev.Hoja, excepcion.Message, StringComparison.Ordinal);
        Assert.Contains(alterada.Celda, excepcion.Message, StringComparison.Ordinal);
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H2.3 — e2e por el CLI real (misma entrada que el operador)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// CA-03: corrida 5-ASE por el CLI REAL (<c>EjecutorCli.Ejecutar</c>, los mismos flags documentados
    /// por <c>--help</c>) a un directorio temporal fresco. Criterios: la hoja <c>ANT EXT-REV</c> coincide
    /// con el manual por bloque/ASE ±0,5; <c>ANTICIPOS USUARIOS</c> y <c>CONSOLIDADO_TOTAL
    /// RECAUDO!L125:L131</c> (M6: L125:L129 referencian, L130 vacía, L131 consumidora) conservan sus
    /// fórmulas y sus valores de plantilla — la app NO los escribe — y el workbook pide recálculo al
    /// abrir (<c>fullCalcOnLoad</c>).
    /// </summary>
    [Theory]
    [InlineData("2026091")]
    [InlineData("2026101")]
    public void Cli_5Ase_AntExtRevIgualAlManual_AnticiposYReferenciasSinEscritura(string periodo)
    {
        var salidaDir = NuevoDirectorio();
        var stdout = new StringWriter();
        var stderr = new StringWriter();

        var codigo = EjecutorCli.Ejecutar(
        [
            "--periodo", periodo,
            "--carpeta", CarpetaInsumos(periodo),
            "--plantilla", PlantillaV4(),
            "--salida", salidaDir,
            "--cinco-ase",
            "--sobrescribir"
        ], stdout, stderr);

        Assert.True(
            codigo == CodigosSalida.Ok || codigo == CodigosSalida.NovedadesEntrega,
            $"El CLI devolvió {codigo}.\nSTDOUT:\n{stdout}\nSTDERR:\n{stderr}");

        var salida = Directory.EnumerateFiles(salidaDir, "*.xlsx", SearchOption.TopDirectoryOnly).Single();

        // ANT EXT-REV ≡ manual por bloque/ASE ±0,5 (destino resuelto por la autoridad de producción).
        var destinos = DestinosDe(Manual(periodo));
        var celdasManual = LeerValoresDeHoja(Manual(periodo), ResolvedorDestinoAntExtRev.Hoja);
        var celdasSalida = LeerValoresDeHoja(salida, ResolvedorDestinoAntExtRev.Hoja);
        Assert.Equal(100, destinos.Count);
        foreach (var destino in destinos)
        {
            Assert.InRange(
                ValorDeCelda(celdasSalida, destino.Celda) - ValorDeCelda(celdasManual, destino.Celda),
                -Tolerancia,
                Tolerancia);
        }

        // Fórmulas protegidas de la hoja: diff vacío contra la plantilla.
        AssertFormulasIdenticas(
            LeerFormulasDeHoja(PlantillaV4(), ResolvedorDestinoAntExtRev.Hoja),
            LeerFormulasDeHoja(salida, ResolvedorDestinoAntExtRev.Hoja));

        // M6: ANTICIPOS USUARIOS y las referencias L125:L131 no se escriben (fórmulas y valores intactos).
        AssertFormulasIdenticas(
            LeerFormulasDeHoja(PlantillaV4(), "ANTICIPOS USUARIOS"),
            LeerFormulasDeHoja(salida, "ANTICIPOS USUARIOS"));
        AssertValoresIdenticos(
            LeerValoresDeHoja(PlantillaV4(), "ANTICIPOS USUARIOS"),
            LeerValoresDeHoja(salida, "ANTICIPOS USUARIOS"));

        var formulasTotal = LeerFormulasDeHoja(salida, "CONSOLIDADO_TOTAL RECAUDO");
        var formulasTotalPlantilla = LeerFormulasDeHoja(PlantillaV4(), "CONSOLIDADO_TOTAL RECAUDO");
        foreach (var celda in new[] { "L125", "L126", "L127", "L128", "L129", "L131" })
        {
            Assert.True(formulasTotal.ContainsKey(celda), $"La celda {celda} de CONSOLIDADO_TOTAL RECAUDO dejó de ser fórmula.");
            Assert.Equal(formulasTotalPlantilla[celda], formulasTotal[celda]);
        }

        Assert.False(formulasTotal.ContainsKey("L130"), "L130 está VACÍA en la plantilla (M6) y la app no debe escribirla.");
        Assert.DoesNotContain("L130", LeerValoresDeHoja(salida, "CONSOLIDADO_TOTAL RECAUDO").Keys);

        Assert.Contains("L131", formulasTotal.Keys);
        Assert.Contains("SUM(L125:L130)", formulasTotal["L131"], StringComparison.OrdinalIgnoreCase);

        Assert.Contains("fullCalcOnLoad=\"1\"", LeerWorkbookXml(salida), StringComparison.Ordinal);
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Utilidades (OpenXml + bytes)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static IReadOnlyList<ResolvedorDestinoAntExtRev.Destino> DestinosDe(string ruta)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        return ResolvedorDestinoAntExtRev.Resolver(documento.WorkbookPart!);
    }

    private static Worksheet Hoja(WorkbookPart parte, string nombre)
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

    private static void AssertFormulasIdenticas(
        IReadOnlyDictionary<string, string> plantilla,
        IReadOnlyDictionary<string, string> salida)
    {
        var diferencias = plantilla.Keys
            .Union(salida.Keys, StringComparer.OrdinalIgnoreCase)
            .Where(celda => !(plantilla.TryGetValue(celda, out var esperada)
                              && salida.TryGetValue(celda, out var observada)
                              && string.Equals(esperada, observada, StringComparison.Ordinal)))
            .Select(celda => $"{celda}: plantilla='{plantilla.GetValueOrDefault(celda)}' salida='{salida.GetValueOrDefault(celda)}'")
            .ToArray();

        Assert.True(
            diferencias.Length == 0,
            "El diff de fórmulas no es vacío (values-only violado):\n" + string.Join("\n", diferencias));
    }

    private static void AssertValoresIdenticos(
        IReadOnlyDictionary<string, string> plantilla,
        IReadOnlyDictionary<string, string> salida)
    {
        var diferencias = plantilla.Keys
            .Union(salida.Keys, StringComparer.OrdinalIgnoreCase)
            .Where(celda => !(plantilla.TryGetValue(celda, out var esperada)
                              && salida.TryGetValue(celda, out var observada)
                              && string.Equals(esperada, observada, StringComparison.Ordinal)))
            .Select(celda => $"{celda}: plantilla='{plantilla.GetValueOrDefault(celda)}' salida='{salida.GetValueOrDefault(celda)}'")
            .ToArray();

        Assert.True(
            diferencias.Length == 0,
            "La app escribió valores en una hoja que no le pertenece:\n" + string.Join("\n", diferencias));
    }

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
        var destino = destinos[hoja.Attribute(Rns + "id")!.Value].Replace("worksheets/", "xl/worksheets/", StringComparison.Ordinal);
        if (!destino.StartsWith("xl/", StringComparison.Ordinal))
        {
            destino = "xl/" + destino;
        }

        return XDocument.Load(zip.GetEntry(destino)!.Open());
    }

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
    /// Fórmulas de una hoja por texto + atributos (<c>t</c>/<c>si</c>/<c>ref</c>): la comparación es
    /// byte-a-byte sobre el TEXTO-<c>&lt;f&gt;</c>, no sobre el <c>&lt;v&gt;</c> cacheado.
    /// </summary>
    private static Dictionary<string, string> LeerFormulasDeHoja(string ruta, string nombreHoja)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var relaciones = XDocument.Load(zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open());
        var hoja = LeerHojaXml(zip, relaciones, nombreHoja);

        return hoja.Descendants(Ns + "c")
            .Where(celda => celda.Attribute("r") is not null && celda.Element(Ns + "f") is not null)
            .ToDictionary(
                celda => (string)celda.Attribute("r")!,
                celda =>
                {
                    var formula = celda.Element(Ns + "f")!;
                    var atributos = string.Join(
                        ",",
                        formula.Attributes().OrderBy(a => a.Name.LocalName)
                            .Select(a => $"{a.Name.LocalName}={a.Value}"));
                    return $"{formula.Value}|{atributos}";
                },
                StringComparer.OrdinalIgnoreCase);
    }

    private static string LeerWorkbookXml(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        using var stream = zip.GetEntry("xl/workbook.xml")!.Open();
        using var lector = new StreamReader(stream);
        return lector.ReadToEnd();
    }
}
