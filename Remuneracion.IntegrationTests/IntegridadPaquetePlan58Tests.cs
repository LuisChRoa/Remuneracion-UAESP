using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 58 (H0/H1/H2) — P0: el workbook generado debe abrir en Excel SIN reparación y sin perder
/// fórmulas.
///
/// Causa raíz medida: <c>INTERVENTORIA!R26</c> es el maestro de un grupo de fórmula compartida
/// (<c>t="shared" ref="R26:R31" si="0"</c>). Phase 0 reescribía el elemento <c>&lt;f&gt;</c> completo
/// y perdía <c>t/ref/si</c>, dejando <c>R27:R31</c> como seguidoras de un grupo sin maestro. Excel
/// detecta el grupo huérfano y "repara" el libro ELIMINANDO las fórmulas (todas en 0 hasta un F9).
///
/// Estos tests fijan, con bytes: (a) la preparación conserva el maestro shared; (b) el escritor
/// rechaza escribir sobre una celda con <c>&lt;f&gt;</c>; (c) el gate de integridad enrojece ante un
/// grupo huérfano y ante una fórmula alterada; (d) el <c>OpenXmlValidator</c> es CIEGO a esta clase
/// de defecto (por eso no puede ser el gate).
/// </summary>
public sealed class IntegridadPaquetePlan58Tests
{
    private const string HojaInterventoria = "INTERVENTORIA";
    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    private const string CeldaTablaJ = "R26";

    // ── (a) La preparación de la base preserva el grupo shared ─────────────────────────────────

    [Fact]
    public void PreparacionBase_Q1_ConservaElMaestroSharedDeLaTablaJ()
    {
        var baseQ1 = BaseDelPeriodo("2026101");

        var formula = LeerFormula(baseQ1, HojaInterventoria, CeldaTablaJ);

        // Q1 deriva el texto de la tabla J a la columna de la quincena (N26) — cambio de TEXTO
        // autorizado por la regla del dominio — pero los atributos shared son de la plantilla.
        Assert.Equal("N26", formula.Texto);
        Assert.Equal("shared", formula.Tipo);
        Assert.Equal("R26:R31", formula.Rango);
        Assert.Equal("0", formula.Indice);
        Assert.Equal(1, ContarMaestrosShared(baseQ1, HojaInterventoria, "0"));
        Assert.Equal(5, ContarSeguidorasShared(baseQ1, HojaInterventoria, "0"));

        // El gate autocontenido no ve ningún grupo huérfano en la base preparada.
        ValidadorIntegridadPaquete.ValidarPaquete(baseQ1);
    }

    [Fact]
    public void PreparacionBase_Q2_ConservaLaFormaSharedOriginalDeLaTablaJ()
    {
        var baseQ2 = BaseDelPeriodo("2026072");

        var formula = LeerFormula(baseQ2, HojaInterventoria, CeldaTablaJ);

        // Q2 = identidad del oficial: mismo texto y mismos atributos, byte a byte.
        Assert.Equal("J15", formula.Texto);
        Assert.Equal("shared", formula.Tipo);
        Assert.Equal("R26:R31", formula.Rango);
        Assert.Equal("0", formula.Indice);
        Assert.Equal(1, ContarMaestrosShared(baseQ2, HojaInterventoria, "0"));

        ValidadorIntegridadPaquete.ValidarPaquete(baseQ2);
    }

    [Fact]
    public void SalidaDelEscritor_SingleAse_NoTieneGruposSharedHuerfanos_YConservaD109()
    {
        var (resultado, leaf) = CasoValidoAse1();
        var dir = Path.Combine(Path.GetTempPath(), "p58-single-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var salida = Path.Combine(dir, "salida.xlsx");

        try
        {
            // El gate corre dentro de GenerarWorkbook; si el paquete no fuera íntegro, esto lanzaría
            // ERR-INTEGRIDAD y el archivo no existiría.
            new OpenXmlPlantillaWriter().GenerarWorkbook(Insumos.Plantilla, salida, resultado, leaf);

            Assert.True(File.Exists(salida), "La generación debió entregar el archivo (paquete íntegro).");
            ValidadorIntegridadPaquete.ValidarPaquete(salida);

            Assert.Equal(1, ContarMaestrosShared(salida, HojaInterventoria, "0"));
            var d109 = LeerFormula(salida, HojaConsolidado, "D109");
            Assert.Equal("SUM(D104:D108)", d109.Texto);
        }
        finally
        {
            BorrarArbol(dir);
        }
    }

    // ── (b) El escritor no puede escribir sobre una celda con <f> ─────────────────────────────

    [Fact]
    public void GuardEscritor_CeldaConFormula_LanzaErrPlantilla_NombrandoHojaYCelda()
    {
        var celda = new Cell
        {
            CellReference = "K26",
            CellFormula = new CellFormula("J15")
        };

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            GuardEscrituraFormula.ExigirSinFormula(celda, HojaInterventoria, "K26", "prueba de guard"));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains(HojaInterventoria, ex.Message, StringComparison.Ordinal);
        Assert.Contains("K26", ex.Message, StringComparison.Ordinal);
        Assert.Contains("values-only", ex.Message, StringComparison.Ordinal);
        Assert.True(GuardEscrituraFormula.EsFormula(celda));
    }

    [Fact]
    public void GuardEscritor_CeldaSinFormula_NoLanzaYPermiteEscribir()
    {
        var celda = new Cell { CellReference = "K26", CellValue = new CellValue("1") };

        GuardEscrituraFormula.ExigirSinFormula(celda, HojaInterventoria, "K26", "prueba de guard");

        Assert.False(GuardEscrituraFormula.EsFormula(celda));
    }

    // ── (c) El gate enrojece: grupo huérfano y fórmula alterada (inyectados) ─────────────────

    [Fact]
    public void Gate_GrupoSharedHuerfanoInjertado_EnrojeceConErrIntegridad()
    {
        var baseQ1 = File.ReadAllBytes(BaseDelPeriodo("2026101"));

        // Injerto del P0 exacto: el maestro R26 pierde t/ref/si y R27:R31 quedan huérfanas.
        var injertado = MutarHojas(
            baseQ1,
            "<f t=\"shared\" ref=\"R26:R31\" si=\"0\">N26</f>",
            "<f>N26</f>");

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            ValidadorIntegridadPaquete.ValidarPaquete(injertado));

        Assert.Equal(CodigoError.Integridad, ex.Codigo);
        Assert.Contains("GRUPO-SHARED", ex.Message, StringComparison.Ordinal);
        Assert.Contains(HojaInterventoria, ex.Message, StringComparison.Ordinal);
        Assert.Contains("si=0", ex.Message, StringComparison.Ordinal);
        Assert.Contains("R27", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Gate_PreparacionDetectaFormulaQuePierdeAtributosShared_ClaseQueHabriaAtrapadoElP0()
    {
        var plantilla = File.ReadAllBytes(Insumos.PlantillaOficial());
        var baseQ1 = File.ReadAllBytes(BaseDelPeriodo("2026101"));

        // Reproduce el P0 en la BASE (no en la salida): el maestro se queda sin t/ref/si.
        var baseRota = MutarHojas(
            baseQ1,
            "<f t=\"shared\" ref=\"R26:R31\" si=\"0\">N26</f>",
            "<f>N26</f>");

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            ValidadorIntegridadPaquete.ValidarPreparacion(plantilla, baseRota));

        Assert.Equal(CodigoError.Integridad, ex.Codigo);
        Assert.Contains("FORMULA-ATRIBUTO-ALTERADO", ex.Message, StringComparison.Ordinal);
        Assert.Contains(HojaInterventoria, ex.Message, StringComparison.Ordinal);
        Assert.Contains("R26", ex.Message, StringComparison.Ordinal);
        Assert.Contains("t='shared'", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Gate_FormulaAlteradaInjertada_EnrojeceNombrandoEsperadoYObtenido()
    {
        var baseQ1 = File.ReadAllBytes(BaseDelPeriodo("2026101"));

        // La misma forma shared, pero con OTRO texto de fórmula (el escritor no puede cambiarlo).
        var alterado = MutarHojas(
            baseQ1,
            "<f t=\"shared\" ref=\"R26:R31\" si=\"0\">N26</f>",
            "<f t=\"shared\" ref=\"R26:R31\" si=\"0\">N25</f>");

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            ValidadorIntegridadPaquete.ValidarEscritura(baseQ1, alterado));

        Assert.Equal(CodigoError.Integridad, ex.Codigo);
        Assert.Contains("FORMULA-ALTERADA", ex.Message, StringComparison.Ordinal);
        Assert.Contains(HojaInterventoria, ex.Message, StringComparison.Ordinal);
        Assert.Contains("esperado 'N26'", ex.Message, StringComparison.Ordinal);
        Assert.Contains("obtenido 'N25'", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Gate_FormulaCreadaPorLaEscritura_Enrojece()
    {
        var baseQ1 = File.ReadAllBytes(BaseDelPeriodo("2026101"));

        var conFormulaNueva = MutarHojas(
            baseQ1,
            "<c r=\"N3\" ",
            "<c r=\"ZZ3\"><f>1+1</f></c><c r=\"N3\" ");

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            ValidadorIntegridadPaquete.ValidarEscritura(baseQ1, conFormulaNueva));

        Assert.Equal(CodigoError.Integridad, ex.Codigo);
        Assert.Contains("FORMULA-CREADA", ex.Message, StringComparison.Ordinal);
        Assert.Contains("ZZ3", ex.Message, StringComparison.Ordinal);
    }

    // ── (c bis) Clases restantes del gate sobre paquetes sintéticos ────────────────────────────

    [Fact]
    public void Gate_CeldaDuplicada_Enrojece()
    {
        var paquete = PaqueteMinimo("<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"A1\"><v>2</v></c></row>");

        var ex = Assert.Throws<CalculoInvalidoException>(() => ValidadorIntegridadPaquete.ValidarPaquete(paquete));

        Assert.Equal(CodigoError.Integridad, ex.Codigo);
        Assert.Contains("CELDA-DUPLICADA", ex.Message, StringComparison.Ordinal);
        Assert.Contains("Hoja1!A1", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Gate_FilaDesordenada_Enrojece()
    {
        var paquete = PaqueteMinimo(
            "<row r=\"2\"><c r=\"A2\"><v>2</v></c></row><row r=\"1\"><c r=\"A1\"><v>1</v></c></row>");

        var ex = Assert.Throws<CalculoInvalidoException>(() => ValidadorIntegridadPaquete.ValidarPaquete(paquete));

        Assert.Equal(CodigoError.Integridad, ex.Codigo);
        Assert.Contains("FILA-DESORDENADA", ex.Message, StringComparison.Ordinal);
        Assert.Contains("Hoja1!A1", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Gate_SharedStringFueraDeRango_Enrojece()
    {
        var paquete = PaqueteMinimo(
            "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>7</v></c></row>",
            sharedStrings: "<sst count=\"1\" uniqueCount=\"1\"><si><t>x</t></si></sst>");

        var ex = Assert.Throws<CalculoInvalidoException>(() => ValidadorIntegridadPaquete.ValidarPaquete(paquete));

        Assert.Equal(CodigoError.Integridad, ex.Codigo);
        Assert.Contains("SHAREDSTRING-FUERA-DE-RANGO", ex.Message, StringComparison.Ordinal);
        Assert.Contains("Hoja1!A1", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Gate_FormulaConHojaInexistente_Enrojece()
    {
        var paquete = PaqueteMinimo("<row r=\"1\"><c r=\"A1\"><f>HojaFantasma!A1</f><v>0</v></c></row>");

        var ex = Assert.Throws<CalculoInvalidoException>(() => ValidadorIntegridadPaquete.ValidarPaquete(paquete));

        Assert.Equal(CodigoError.Integridad, ex.Codigo);
        Assert.Contains("HOJA-INEXISTENTE", ex.Message, StringComparison.Ordinal);
        Assert.Contains("HojaFantasma", ex.Message, StringComparison.Ordinal);
    }

    // ── (d) El OpenXmlValidator es CIEGO al grupo huérfano (R-VALIDATOR-CIEGO) ─────────────────

    [Fact]
    public void OpenXmlValidator_NoDetectaElGrupoSharedHuerfano()
    {
        var baseQ1 = File.ReadAllBytes(BaseDelPeriodo("2026101"));
        var injertado = MutarHojas(
            baseQ1,
            "<f t=\"shared\" ref=\"R26:R31\" si=\"0\">N26</f>",
            "<f>N26</f>");

        var erroresSano = ValidarConOpenXmlValidator(baseQ1);
        var erroresRoto = ValidarConOpenXmlValidator(injertado);

        // El SDK no distingue el paquete roto del sano: no hay error nuevo y ninguno menciona la
        // fórmula compartida ni la celda maestra. Por eso el gate es propio (Plan 58, §2-D2).
        Assert.Equal(erroresSano.Count, erroresRoto.Count);
        Assert.DoesNotContain(erroresRoto, e =>
            e.Contains("shared", StringComparison.OrdinalIgnoreCase)
            || e.Contains(CeldaTablaJ, StringComparison.OrdinalIgnoreCase));
    }

    // ── Helpers ────────────────────────────────────────────────────────────────────────────────

    private static (ResultadoRemuneracion Resultado, WorkbookLeafInputs Leaf) CasoValidoAse1()
    {
        var reader = new ExcelDataReaderRecaudoReader();
        var r1 = reader.LeerR1(Insumos.R1(1));
        var r2 = reader.LeerR2(Insumos.R2(1));
        var r4 = reader.LeerR4(Insumos.R4(1));
        var ase = Insumos.Ase(1);
        var periodo = Insumos.Periodo();
        var resultado = new CalculoRemuneracion().CalcularConsolidado(periodo, [(ase, r1, r2, r4)]);
        var leaf = new ExcelDataReaderWorkbookLeafInputReader()
            .LeerLeafInputs(ase, periodo, Insumos.R1(1), Insumos.R2(1), Insumos.R4(1));
        return (resultado, leaf);
    }

    /// <summary>
    /// Base del período producida por la ruta de producción (Phase 0 desde la plantilla oficial v2),
    /// en una carpeta temp fresca por llamada.
    /// </summary>
    private static string BaseDelPeriodo(string periodo)
    {
        var dir = Path.Combine(Path.GetTempPath(), $"p58-base-{periodo}-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var salida = Path.Combine(dir, $"base-{periodo}.xlsx");
        PreparadorBasePipeline.Ejecutar(Insumos.PlantillaOficial(), Periodo.Parse(periodo), salida);
        return salida;
    }

    private static (string? Texto, string? Tipo, string? Rango, string? Indice) LeerFormula(
        string ruta,
        string hoja,
        string celda)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var worksheet = Hoja(documento, hoja);
        var c = worksheet.Descendants<Cell>()
            .FirstOrDefault(x => string.Equals(x.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe la celda '{hoja}!{celda}'.");

        var formula = c.CellFormula
            ?? throw new InvalidOperationException($"La celda '{hoja}!{celda}' no tiene <f>.");

        return (
            formula.Text,
            formula.FormulaType?.InnerText,
            formula.Reference?.Value,
            formula.SharedIndex?.Value.ToString(CultureInfo.InvariantCulture));
    }

    private static int ContarMaestrosShared(string ruta, string hoja, string indice)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        return Hoja(documento, hoja).Descendants<CellFormula>()
            .Count(f => string.Equals(f.SharedIndex?.Value.ToString(CultureInfo.InvariantCulture), indice, StringComparison.Ordinal)
                && string.Equals(f.FormulaType?.InnerText, "shared", StringComparison.Ordinal)
                && !string.IsNullOrEmpty(f.Reference?.Value));
    }

    private static int ContarSeguidorasShared(string ruta, string hoja, string indice)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        return Hoja(documento, hoja).Descendants<CellFormula>()
            .Count(f => string.Equals(f.SharedIndex?.Value.ToString(CultureInfo.InvariantCulture), indice, StringComparison.Ordinal)
                && string.Equals(f.FormulaType?.InnerText, "shared", StringComparison.Ordinal)
                && string.IsNullOrEmpty(f.Reference?.Value));
    }

    private static Worksheet Hoja(SpreadsheetDocument documento, string hoja)
    {
        var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null");
        var sheet = workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, hoja, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"No existe la hoja '{hoja}'.");
        var parte = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return parte.Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
    }

    private static List<string> ValidarConOpenXmlValidator(byte[] paquete)
    {
        using var memoria = new MemoryStream(paquete, writable: false);
        using var documento = SpreadsheetDocument.Open(memoria, false);
        return new OpenXmlValidator().Validate(documento).Select(e => e.Description).ToList();
    }

    /// <summary>
    /// Aplica una sustitución de texto a TODAS las partes de hoja del paquete y devuelve el zip nuevo
    /// en memoria (la plantilla/base en disco no se toca).
    /// </summary>
    private static byte[] MutarHojas(byte[] paquete, string buscar, string reemplazar)
    {
        using var entrada = new MemoryStream(paquete, writable: false);
        using var zipEntrada = new ZipArchive(entrada, ZipArchiveMode.Read);
        using var salida = new MemoryStream();
        using (var zipSalida = new ZipArchive(salida, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var item in zipEntrada.Entries)
            {
                using var origen = item.Open();
                using var memoriaParte = new MemoryStream();
                origen.CopyTo(memoriaParte);
                var bytes = memoriaParte.ToArray();

                if (item.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase))
                {
                    var texto = Encoding.UTF8.GetString(bytes);
                    if (texto.Contains(buscar, StringComparison.Ordinal))
                    {
                        bytes = new UTF8Encoding(false).GetBytes(texto.Replace(buscar, reemplazar, StringComparison.Ordinal));
                    }
                }

                var nueva = zipSalida.CreateEntry(item.FullName, CompressionLevel.Optimal);
                using var destino = nueva.Open();
                destino.Write(bytes);
            }
        }

        return salida.ToArray();
    }

    /// <summary>Workbook mínimo y VÁLIDO para el parser del gate (1 hoja, sin estilos).</summary>
    private static byte[] PaqueteMinimo(string filas, string? sharedStrings = null)
    {
        using var salida = new MemoryStream();
        using (var zip = new ZipArchive(salida, ZipArchiveMode.Create, leaveOpen: true))
        {
            Escribir(zip, "xl/workbook.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<sheets><sheet name=\"Hoja1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");

            Escribir(zip, "xl/_rels/workbook.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" " +
                "Target=\"worksheets/sheet1.xml\"/></Relationships>");

            Escribir(zip, "xl/worksheets/sheet1.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                $"<sheetData>{filas}</sheetData></worksheet>");

            if (sharedStrings is not null)
            {
                Escribir(zip, "xl/sharedStrings.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    $"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{sharedStrings}</sst>");
            }
        }

        return salida.ToArray();
    }

    private static void Escribir(ZipArchive zip, string nombre, string contenido)
    {
        var entrada = zip.CreateEntry(nombre, CompressionLevel.Optimal);
        using var stream = entrada.Open();
        var bytes = new UTF8Encoding(false).GetBytes(contenido);
        stream.Write(bytes);
    }

    private static void BorrarArbol(string dir)
    {
        try
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, recursive: true);
            }
        }
        catch (IOException)
        {
            // Best-effort: un handle de antivirus no debe enmascarar el resultado del test.
        }
    }
}
