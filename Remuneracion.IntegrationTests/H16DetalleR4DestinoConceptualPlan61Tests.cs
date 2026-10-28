using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Cli;
using Remuneracion.Infrastructure.Excel;
using Xunit;
using Xunit.Abstractions;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// PLAN 61 — H1.6 (D9 ratificada): el DESTINO de las columnas del desglose-detalle se resuelve por
/// la IDENTIDAD CONCEPTUAL de la plantilla, no por una lista congelada de pares encabezado→letra.
///
/// <para><b>El defecto que cierra.</b> <c>WorkbookLeafCellMapDetalleR2R4.ColumnasR4</c> fijaba 13
/// pares y no incluía <c>SERVICIO ESPECIALES</c> ni <c>Ajuste Recargo Aseo</c>: la plantilla declara
/// 15 columnas de valor en ese bloque, así que un componente con monto y sin entrada en la lista se
/// DESCARTABA EN SILENCIO (3 de los 10 R4 de septiembre bloqueaban por el gate de valores y 3
/// perdían centavos dentro de la tolerancia). La escritura y el verificador no compartían autoridad.</para>
///
/// <para><b>Lo que fija este fixture.</b></para>
/// <list type="bullet">
/// <item>H161: invariante de alcanzabilidad — todo encabezado que la plantilla v4 declara en el
/// bloque de cada ASE resuelve a SU propia columna, y la clase de alias de Especiales resuelve a una
/// sola (la tabla es única: <see cref="ResolvedorConciliacionEmpresa.EncabezadosEspeciales"/>).</item>
/// <item>H162 (injerto): un componente sintético con monto y SIN destino en la plantilla aborta en
/// duro nombrando ASE · sección · concepto · encabezado, y no deja archivo. Nunca se escribe 0 ni se
/// ignora.</item>
/// <item>H163: corridas reales 2026092 (ASE 2 y ASE 4) escriben <c>Ajuste Recargo Aseo</c> en Q y
/// <c>SERVICIO ESPECIALES</c> en P, con los montos del administrativo: <c>CONSOLIDADO!AC67 = 279,03</c>
/// (Q+R del cierre del ASE 2) y <c>CONSOLIDADO!AB69 = −49.900,44</c> (P del cierre del ASE 4).</item>
/// <item>H164 (la clase, no el caso): los tres archivos que perdían centavos en silencio —
/// Julio-1 ASE5 (+0,02), Octubre-1 ASE2 (+0,16) y Septiembre-2 ASE5 (−0,09) — escriben ahora su
/// componente en la columna que la plantilla declara.</item>
/// </list>
/// </summary>
public sealed class H16DetalleR4DestinoConceptualPlan61Tests
{
    private readonly ITestOutputHelper _salida;

    public H16DetalleR4DestinoConceptualPlan61Tests(ITestOutputHelper salida) => _salida = salida;
    static H16DetalleR4DestinoConceptualPlan61Tests()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    private const string HojaR4 = "Reversion Pagos R4";    private const string HojaR2 = "Rem. Anticipos R2";
    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    private const decimal Centavo = 0.005m;

    private static string Raiz => RaizRepositorio();

    private static string RutaPlantillaV4 => Path.Combine(
        Raiz, "Docs", "Plantilla estandarizada", "v4", "Formato Remuneracion 2026xx-x Total v4.xlsx");

    private static string CarpetaSeptiembre2 => Path.Combine(Raiz, "Docs", "Prueba Septiembre-2", "Insumos");
    private static string CarpetaJulio1 => Path.Combine(Raiz, "Docs", "Prueba Julio-1", "Insumos");
    private static string CarpetaOctubre1 => Path.Combine(Raiz, "Docs", "Prueba Octubre-1", "Insumos");

    /// <summary>
    /// Corrida de CINCO ASE de 2026092, compartida por los tests de escritura. La quincena 2 exige el
    /// camino multi-ASE (el camino de un solo ASE declara no soportar SALDOS POR NOTA / RETRIBUCIÓN
    /// NEGATIVA), así que las verificaciones de ASE 2/ASE 4 y del caso de centavos de septiembre se
    /// hacen sobre esta única generación.
    /// </summary>
    private static readonly Lazy<Corrida> Generacion2026092 = new(() =>
        EjecutarPeriodo("2026092", CarpetaSeptiembre2, aseId: null));

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // H161 — invariante conceptual: alcanzabilidad + alias único
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void H161_PlantillaV4_TodoEncabezadoDeValorDelBloque_EsAlcanzable(int aseId)
    {
        foreach (var hoja in new DelimitadorBloquesHoja.HojaDelimitable[] { DelimitadorBloquesHoja.R4, DelimitadorBloquesHoja.R2 })
        {
            using var workbook = SpreadsheetDocument.Open(RutaPlantillaV4, false);
            var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var worksheet = Hoja(workbookPart, hoja.Nombre);
            var identidad = $"hoja '{hoja.Nombre}' · ASE {aseId}";
            var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hoja, aseId, worksheet);
            var columnas = ResolvedorDestinoDetalleR2R4.ColumnasDeclaradas(
                workbookPart, worksheet, bloque.EncabezadoColumnas, identidad);
            var indice = ResolvedorDestinoDetalleR2R4.IndicePorEtiqueta(columnas, identidad);

            Assert.NotEmpty(columnas);
            foreach (var (etiqueta, columna) in columnas)
            {
                Assert.Equal(columna, ResolvedorDestinoDetalleR2R4.ColumnaDe(indice, etiqueta));
            }

            // El alias de Especiales es UNA sola clase de equivalencia: cualquier miembro resuelve a la
            // columna que el bloque declare para ese componente (o a ninguna si el bloque no lo trae).
            var deEspeciales = columnas
                .Where(c => ResolvedorDestinoDetalleR2R4.PerteneceALaClaseDeAlias(c.Etiqueta))
                .Select(c => c.Columna)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            Assert.True(deEspeciales.Count <= 1, $"{identidad}: {deEspeciales.Count} columnas distintas para la clase de Especiales.");
            foreach (var alias in ResolvedorConciliacionEmpresa.EncabezadosEspeciales)
            {
                Assert.Equal(
                    deEspeciales.Count == 0 ? null : deEspeciales[0],
                    ResolvedorDestinoDetalleR2R4.ColumnaDe(indice, alias));
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // H162 — injerto: componente con monto y sin destino ⇒ fallo duro (no-pérdida)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H162_R4_ComponenteConMontoSinDestino_FallaEnDuro_YNoDejaArchivo()
    {
        const string componente = "Componente SINTETICO INJERTADO";
        const decimal monto = 12345.67m;

        var carpeta = CopiarCarpetaInsumos(CarpetaSeptiembre2);
        var rutaR4Ase4 = LocalizarPorPrefijo(Path.Combine(carpeta, "4-Bogota Limpia"), "ReversiónPorComponente");
        InjertarComponenteEnFuenteR4(rutaR4Ase4, componente, monto);

        var salida = Path.Combine(Path.GetTempPath(), $"p61-h16-injerto-{Environment.ProcessId}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(salida);

        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var codigo = EjecutorCli.Ejecutar(
            ["--periodo", "2026092", "--carpeta", carpeta, "--plantilla", RutaPlantillaV4,
                "--salida", salida, "--cinco-ase", "--sobrescribir"],
            stdout,
            stderr);

        var log = stdout + stderr.ToString();
        _salida.WriteLine(log);
        Assert.NotEqual(0, codigo);
        Assert.Contains(componente, log, StringComparison.Ordinal);
        Assert.Contains("ASE 4", log, StringComparison.Ordinal);
        Assert.Contains("ERR-PLANTILLA", log, StringComparison.Ordinal);
        Assert.Contains("12345", log, StringComparison.Ordinal);

        // La identidad del valor viaja completa en el mensaje: la fila se nombra por sus rótulos.
        Assert.Contains("concepto 'Vlr Servicio'", log, StringComparison.Ordinal);

        // No-pérdida también significa no-invención: el archivo no se entrega (ni con nombre limpio ni
        // `(sin validar)`), y el parcial se borra.
        Assert.Empty(Directory.EnumerateFiles(salida, "*.xlsx", SearchOption.AllDirectories));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // H163 — corridas reales 2026092: P y Q escritas en la columna declarada
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H163_2026092_Ase2_EscribeAjusteRecargoAseo_EnQ_YCierraAC67()
    {
        var corrida = Generacion2026092.Value;
        Assert.Equal(0, corrida.Codigo);

        using var workbook = SpreadsheetDocument.Open(corrida.RutaSalida, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, HojaR4);
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, DelimitadorBloquesHoja.R4, 2, worksheet);
        var columnas = ResolvedorDestinoDetalleR2R4.ColumnasDeclaradas(
            workbookPart, worksheet, bloque.EncabezadoColumnas, $"hoja '{HojaR4}' · ASE 2");
        var indice = ResolvedorDestinoDetalleR2R4.IndicePorEtiqueta(columnas, $"hoja '{HojaR4}' · ASE 2");

        var columnaAjuste = ResolvedorDestinoDetalleR2R4.ColumnaDe(indice, "Ajuste Recargo Aseo");
        var columnaOtros = ResolvedorDestinoDetalleR2R4.ColumnaDe(indice, "Otros");
        Assert.NotNull(columnaAjuste);
        Assert.NotNull(columnaOtros);

        var ajuste = ValorDeCelda(workbookPart, worksheet, columnaAjuste! + bloque.GranTotal.ToString(CultureInfo.InvariantCulture));
        var otros = ValorDeCelda(workbookPart, worksheet, columnaOtros! + bloque.GranTotal.ToString(CultureInfo.InvariantCulture));
        Assert.Equal(-10.16m, ajuste);
        Assert.Equal(289.19m, otros);

        // CONSOLIDADO!AC67 = Q+R del cierre del ASE 2 = 279,03 (el valor del administrativo 202609-2).
        Assert.Equal(279.03m, ajuste + otros);
    }

    [Fact]
    public void H163_2026092_Ase4_EscribeServicioEspeciales_EnP_YCierraAB69()
    {
        var corrida = Generacion2026092.Value;
        Assert.Equal(0, corrida.Codigo);

        using var workbook = SpreadsheetDocument.Open(corrida.RutaSalida, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, HojaR4);
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, DelimitadorBloquesHoja.R4, 4, worksheet);
        var columnas = ResolvedorDestinoDetalleR2R4.ColumnasDeclaradas(
            workbookPart, worksheet, bloque.EncabezadoColumnas, $"hoja '{HojaR4}' · ASE 4");
        var indice = ResolvedorDestinoDetalleR2R4.IndicePorEtiqueta(columnas, $"hoja '{HojaR4}' · ASE 4");

        var columnaEspeciales = ResolvedorDestinoDetalleR2R4.ColumnaDe(indice, "SERVICIO ESPECIALES");
        Assert.NotNull(columnaEspeciales);

        var especiales = ValorDeCelda(workbookPart, worksheet, columnaEspeciales! + bloque.GranTotal.ToString(CultureInfo.InvariantCulture));
        Assert.Equal(-49900.44m, especiales);

        // CONSOLIDADO!AB69 = N69 = P del cierre del ASE 4 (el valor del administrativo 202609-2). La
        // fórmula es SHARED en la plantilla: se expande contra su maestro para afirmar la identidad.
        Assert.Equal("N69", FormulaEfectivaDeCelda(workbookPart, HojaConsolidado, "AB69"));
        Assert.Contains($"'{HojaR4}'!P", FormulaEfectivaDeCelda(workbookPart, HojaConsolidado, "N69"), StringComparison.OrdinalIgnoreCase);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // H164 — la clase completa: los tres archivos que perdían centavos en silencio
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H164_Julio1_Ase5_EscribeAjusteRecargoAseo_EnLaColumnaDeclarada()
    {
        var corrida = EjecutarPeriodo("2026071", CarpetaJulio1, aseId: 5);
        Assert.Equal(0, corrida.Codigo);
        AssertAjusteRecargoAseoEscrito(corrida.RutaSalida, aseId: 5, esperado: 0.02m);
    }

    [Fact]
    public void H164_Octubre1_Ase2_EscribeAjusteRecargoAseo_EnLaColumnaDeclarada()
    {
        var corrida = EjecutarPeriodo("2026101", CarpetaOctubre1, aseId: 2);
        Assert.Equal(0, corrida.Codigo);
        AssertAjusteRecargoAseoEscrito(corrida.RutaSalida, aseId: 2, esperado: 0.16m);
    }

    [Fact]
    public void H164_Septiembre2_Ase5_EscribeAjusteRecargoAseo_EnLaColumnaDeclarada()
    {
        var corrida = Generacion2026092.Value;
        Assert.Equal(0, corrida.Codigo);
        AssertAjusteRecargoAseoEscrito(corrida.RutaSalida, aseId: 5, esperado: -0.09m);
    }

    /// <summary>
    /// El componente <c>Ajuste Recargo Aseo</c> de la fuente se escribe en la columna que la plantilla
    /// declara (Q en la v4) en la fila de cierre del bloque del ASE. Antes ese monto se descartaba y
    /// solo se notaba por el Δ de centavos del <c>Otros</c> del CONSOLIDADO.
    /// </summary>
    private static void AssertAjusteRecargoAseoEscrito(string rutaSalida, int aseId, decimal esperado)
    {
        using var workbook = SpreadsheetDocument.Open(rutaSalida, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, HojaR4);
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, DelimitadorBloquesHoja.R4, aseId, worksheet);
        var columnas = ResolvedorDestinoDetalleR2R4.ColumnasDeclaradas(
            workbookPart, worksheet, bloque.EncabezadoColumnas, $"hoja '{HojaR4}' · ASE {aseId}");
        var indice = ResolvedorDestinoDetalleR2R4.IndicePorEtiqueta(columnas, $"hoja '{HojaR4}' · ASE {aseId}");

        var columna = ResolvedorDestinoDetalleR2R4.ColumnaDe(indice, "Ajuste Recargo Aseo");
        Assert.NotNull(columna);

        var escrito = ValorDeCelda(
            workbookPart,
            worksheet,
            columna! + bloque.GranTotal.ToString(CultureInfo.InvariantCulture));
        Assert.Equal(esperado, escrito);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Helpers
    // ════════════════════════════════════════════════════════════════════════════════════════════

    private sealed record Corrida(int Codigo, string RutaSalida, string Log);

    private static Corrida EjecutarPeriodo(string periodo, string carpeta, int? aseId)
    {
        var salida = Path.Combine(Path.GetTempPath(), $"p61-h16-{periodo}-{aseId?.ToString(CultureInfo.InvariantCulture) ?? "5ase"}-{Environment.ProcessId}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(salida);

        var modo = aseId is null
            ? new[] { "--cinco-ase" }
            : ["--ase", aseId.Value.ToString(CultureInfo.InvariantCulture)];

        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var codigo = EjecutorCli.Ejecutar(
            ["--periodo", periodo, "--carpeta", carpeta, "--plantilla", RutaPlantillaV4,
                "--salida", salida, .. modo, "--sobrescribir"],
            stdout,
            stderr);

        var archivo = Directory.EnumerateFiles(salida, "*.xlsx", SearchOption.TopDirectoryOnly).FirstOrDefault();
        return new Corrida(codigo, archivo ?? string.Empty, stdout + stderr.ToString());
    }

    private static string CopiarCarpetaInsumos(string origen)
    {
        var destino = Path.Combine(Path.GetTempPath(), $"p61-h16-copia-{Environment.ProcessId}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(destino);
        foreach (var archivo in Directory.EnumerateFiles(origen, "*", SearchOption.AllDirectories))
        {
            var relativo = Path.GetRelativePath(origen, archivo);
            var objetivo = Path.Combine(destino, relativo);
            Directory.CreateDirectory(Path.GetDirectoryName(objetivo)!);
            File.Copy(archivo, objetivo);
        }

        return destino;
    }

    private static string LocalizarPorPrefijo(string carpeta, string prefijo) =>
        Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .First(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Injerta en la COPIA de la fuente R4 un encabezado de componente que la plantilla NO declara, con
    /// un monto distinto de cero en la fila de detalle <c>Vlr Servicio</c> y en la fila de cierre
    /// <c>Total</c> del bloque. Es el escenario que antes se descartaba en silencio.
    /// </summary>
    private static void InjertarComponenteEnFuenteR4(string ruta, string encabezado, decimal monto)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, "Sheet1");
        var filas = worksheet.Elements<SheetData>().SelectMany(tabla => tabla.Elements<Row>())
            .Where(f => f.RowIndex is not null)
            .OrderBy(f => f.RowIndex!.Value)
            .ToList();

        var filaVlrServicio = filas.First(f => Texto(workbookPart, f, "C").Equals("Vlr Servicio", StringComparison.OrdinalIgnoreCase));
        var indiceEncabezado = filas.IndexOf(filaVlrServicio) - 1;
        while (indiceEncabezado > 0 && !FilaTieneTextoDesde(filaEncabezado: filas[indiceEncabezado], desde: 3, workbookPart))
        {
            indiceEncabezado--;
        }

        var filaEncabezado = filas[indiceEncabezado];

        var filaTotal = filas.First(f =>
            Texto(workbookPart, f, "A").Equals("Total", StringComparison.OrdinalIgnoreCase)
            && Texto(workbookPart, f, "B").Length == 0);

        // La primera columna libre a la derecha de la última columna con encabezado.
        var ultima = filaEncabezado.Elements<Cell>().Select(Indice).Where(i => i >= 3).Max();
        var destino = ultima + 1;

        AgregarTexto(filaEncabezado, destino, encabezado);
        AgregarNumero(filaVlrServicio, destino, monto);
        AgregarNumero(filaTotal, destino, monto);
        worksheet.Save();
        workbookPart.Workbook?.Save();
    }

    private static bool FilaTieneTextoDesde(Row filaEncabezado, int desde, WorkbookPart workbookPart) =>
        filaEncabezado.Elements<Cell>()
            .Any(c => Indice(c) > desde && Texto(workbookPart, c).Trim().Length > 0);

    private static int Indice(Cell celda)    {
        var letras = new string((celda.CellReference?.Value ?? "A1").Where(char.IsLetter).ToArray());
        var indice = 0;
        foreach (var letra in letras)
        {
            indice = (indice * 26) + (letra - 'A' + 1);
        }

        return indice;
    }

    private static string Letras(int indice)
    {
        var texto = string.Empty;
        while (indice > 0)
        {
            var resto = (indice - 1) % 26;
            texto = (char)('A' + resto) + texto;
            indice = (indice - resto - 1) / 26;
        }

        return texto;
    }

    private static void AgregarTexto(Row fila, int columna, string texto)
    {
        var referencia = Letras(columna) + fila.RowIndex!.Value.ToString(CultureInfo.InvariantCulture);
        fila.Append(new Cell
        {
            CellReference = referencia,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text(texto)),
        });
        fila.Spans = null;
    }

    private static void AgregarNumero(Row fila, int columna, decimal valor)
    {
        var referencia = Letras(columna) + fila.RowIndex!.Value.ToString(CultureInfo.InvariantCulture);
        fila.Append(new Cell
        {
            CellReference = referencia,
            CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture)),
        });
        fila.Spans = null;
    }

    private static string Texto(WorkbookPart workbookPart, Row fila, string columna) =>
        fila.Elements<Cell>()
            .Where(c => string.Equals(Letras(Indice(c)), columna, StringComparison.OrdinalIgnoreCase))
            .Select(c => Texto(workbookPart, c))
            .FirstOrDefault() ?? string.Empty;

    private static string Texto(WorkbookPart workbookPart, Cell celda)
    {
        if (celda.DataType?.Value == CellValues.SharedString
            && int.TryParse(celda.CellValue?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice))
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(indice)?.InnerText ?? string.Empty;
        }

        return celda.InlineString?.InnerText ?? celda.CellValue?.InnerText ?? string.Empty;
    }

    private static decimal ValorDeCelda(WorkbookPart workbookPart, Worksheet worksheet, string referencia)
    {
        var celda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        return celda?.CellValue is null
            ? 0m
            : decimal.Parse(celda.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Fórmula EFECTIVA de una celda: si es una fórmula compartida (la plantilla usa <c>t="shared"</c>
    /// en los cuadros-resumen), se expande el texto del maestro con el desplazamiento de fila/columna
    /// de la celda, igual que hace Excel. Sin esto, la celda compartida declara texto vacío.
    /// </summary>
    private static string FormulaEfectivaDeCelda(WorkbookPart workbookPart, string hoja, string referencia)
    {
        var worksheet = Hoja(workbookPart, hoja);
        var celdas = worksheet.Descendants<Cell>().Where(c => c.CellReference?.Value is not null).ToList();
        var celda = celdas.FirstOrDefault(c => string.Equals(c.CellReference!.Value, referencia, StringComparison.OrdinalIgnoreCase));
        var formula = celda?.CellFormula;
        if (formula is null)
        {
            return string.Empty;
        }

        if (!string.IsNullOrEmpty(formula.Text))
        {
            return formula.Text;
        }

        var indiceCompartido = formula.SharedIndex?.Value;
        if (indiceCompartido is null)
        {
            return string.Empty;
        }

        var maestro = celdas.FirstOrDefault(c =>
            c.CellFormula?.SharedIndex?.Value == indiceCompartido && !string.IsNullOrEmpty(c.CellFormula.Text));
        if (maestro is null)
        {
            return string.Empty;
        }

        var deltaFila = Fila(celda!) - Fila(maestro);
        var deltaColumna = Columna(celda!) - Columna(maestro);
        return System.Text.RegularExpressions.Regex.Replace(
            maestro.CellFormula!.Text,
            "[A-Z]+[0-9]+",
            coincidencia => DesplazarReferencia(coincidencia.Value, deltaFila, deltaColumna));
    }

    private static string DesplazarReferencia(string referencia, int deltaFila, int deltaColumna)
    {
        var letras = new string(referencia.Where(char.IsLetter).ToArray());
        var fila = int.Parse(new string(referencia.Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture);
        var columna = 0;
        foreach (var letra in letras)
        {
            columna = (columna * 26) + (letra - 'A' + 1);
        }

        return Letras(columna + deltaColumna) + (fila + deltaFila).ToString(CultureInfo.InvariantCulture);
    }

    private static int Fila(Cell celda) =>
        int.Parse(new string((celda.CellReference!.Value!).Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture);

    private static int Columna(Cell celda) => Indice(celda);

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var hoja = workbookPart.Workbook?.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"El workbook no trae la hoja '{nombre}'.");
        return ((WorksheetPart)workbookPart.GetPartById(hoja.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene WorksheetPart válido.");
    }

    private static string RaizRepositorio()
    {
        var directorio = new DirectoryInfo(AppContext.BaseDirectory);
        while (directorio is not null)
        {
            if (File.Exists(Path.Combine(directorio.FullName, "AGENTS.md")))
            {
                return directorio.FullName;
            }

            directorio = directorio.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }
}
