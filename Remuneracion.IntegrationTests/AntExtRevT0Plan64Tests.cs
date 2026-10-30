using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ExcelDataReader;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 64 (H0 — T0.1..T0.7). Verificación READ-ONLY por bytes de los seis frentes del §2.1 del
/// plan 64:
/// <list type="bullet">
/// <item>(a) geometría v4 de la hoja <c>ANT EXT-REV</c> (bloque superior, 5 etiquetas, rejilla de
/// valores);</item>
/// <item>(b) <c>ANTICIPOS USUARIOS</c> 100 %-fórmulas y su consumidora
/// <c>CONSOLIDADO_TOTAL RECAUDO!L125:L130</c>;</item>
/// <item>(c) estructura <c>Sheet1</c> de los 30 <c>RecaudosReversados_*.xlsx</c>;</item>
/// <item>(d) universo de fuentes 30/30 + paridad numérica del §1.4 re-derivada;</item>
/// <item>(e) divergencia docx-vs-encargo (<c>D-10</c>);</item>
/// <item>(f) huella v4 antes/después.</item>
/// </list>
///
/// <para>
/// <b>Método (D-H0):</b> BCL puro (<see cref="ZipArchive"/> + <see cref="XDocument"/>) sobre las
/// plantillas y los 4 manuales —rutas EXPLÍCITAS, nunca <c>Insumos.Raiz()</c>— y ExcelDataReader
/// SOLO-lectura sobre COPIAS temporales para los 30 insumos. Ningún .xlsx se abre en escritura y
/// ningún Excel debe estar abierto durante la corrida.
/// </para>
///
/// <para>
/// <b>Regla honesta:</b> cada assert refleja el BYTE medido. Los números citados por el §1 del plan
/// se contrastan aquí; cuando el disco difiere, el assert refleja el DISCO y el hallazgo se declara
/// en <c>Docs/T0-Plan64-Ant-Ext-Rev.md</c> (el §1 no se convierte en contrato).
/// </para>
/// </summary>
public sealed class AntExtRevT0Plan64Tests
{
    private const decimal Tolerancia = 0.5m;

    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Prns = "http://schemas.openxmlformats.org/package/2006/relationships";

    private const string HojaAntExtRev = "ANT EXT-REV";
    private const string HojaAnticipos = "ANTICIPOS USUARIOS";
    private const string HojaConsolidadoTotal = "CONSOLIDADO_TOTAL RECAUDO";

    /// <summary>Marcadores del insumo (texto exacto medido en los 30 archivos).</summary>
    private const string MarcadorBloque2 = "Pagos Pendientes por Aplicar";
    private const string MarcadorBloque3 = "Resumen Recaudo Aplicado Por Servicio";

    /// <summary>
    /// Etiquetas de bloque del template (columna B), en el orden del archivo. Son DATO leído por el
    /// H0: no se usan como contrato posicional.
    /// </summary>
    private static readonly string[] EtiquetasBloque =
    [
        "BANCO DE OCCIDENTE", "EAAB RECIPROCIDAD", "EAAB CIUDAD LIMPIA", "ENEL", "ENERBIT"
    ];

    /// <summary>Catálogo del §1.2 (etiqueta de bloque ← rótulo de fuente del bloque 2), congelado por T0.</summary>
    private static readonly (string EtiquetaBloque, string Fuente)[] CatalogoBloqueFuente =
    [
        ("BANCO DE OCCIDENTE", "OCCIDENTE"),
        ("EAAB RECIPROCIDAD", "NUEVO ESQUEMA"),
        ("ENEL", "ENEL")
    ];

    /// <summary>Conceptos del §1.2 (rótulo de concepto → fila de valor / columna de reporte), congelados por T0.</summary>
    private static readonly (string Concepto, string Fila, string Reporte)[] CatalogoConceptos =
    [
        ("1-APLICADOS A FACTURACION", "APLICADO", "REPORTE 3"),
        ("2-SALDOS A FAVOR GENERADOS", "ANTICIPOS", "REPORTE 3"),
        ("9-REVERSION PAGOS A COMPONENTES", "APLICADO", "REPORTE 4"),
        ("9-REVERSION SALDOS A FAVOR", "ANTICIPOS", "REPORTE 4")
    ];

    /// <summary>Fuentes que el encargo declara ausentes (HUE-1): 0 por ausencia, nunca mapeo inventado.</summary>
    private static readonly string[] FuentesDeclaradasAusentes = ["EAAB CIUDAD LIMPIA", "ENERBIT"];

    static AntExtRevT0Plan64Tests()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
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

    /// <summary>Los 6 períodos con su carpeta explícita de insumos (una entrada por período).</summary>
    private static readonly (string Periodo, string Carpeta)[] Periodos =
    [
        ("2026071", "Prueba Julio-1"),
        ("2026072", "Prueba Julio-2"),
        ("2026082", "Prueba Agosto-2"),
        ("2026091", "Prueba Septiembre 1"),
        ("2026092", "Prueba Septiembre-2"),
        ("2026101", "Prueba Octubre-1")
    ];

    private static string CarpetaInsumos(string periodo) =>
        Ruta("Docs", Periodos.Single(p => p.Periodo == periodo).Carpeta, "Insumos");

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia"
    ];

    private static string CarpetaAse(string periodo, int aseId) =>
        Path.Combine(CarpetaInsumos(periodo), CarpetasAse[aseId - 1]);

    private static string Insumo(string periodo, int aseId)
    {
        var carpeta = CarpetaAse(periodo, aseId);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta del ASE{aseId} en {periodo}: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .SingleOrDefault(f => Path.GetFileName(f).StartsWith("RecaudosReversados_", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta un único RecaudosReversados_*.xlsx en {carpeta}");
        return archivo!;
    }

    /// <summary>Los 4 manuales del administrativo citados por el encargo (rutas explícitas).</summary>
    private static readonly (string Periodo, string Ruta)[] Manuales =
    [
        ("2026072", Path.Combine("Docs", "Prueba Julio-2", "Resultado", "Resultado Administrativo con plantilla estandar v2", "Formato Remuneracion 202607-2 Total.xlsx")),
        ("2026091", Path.Combine("Docs", "Prueba Septiembre 1", "Resultados", "Administrativo", "Remuneracion 202609-1 Total.xlsx")),
        ("2026092", Path.Combine("Docs", "Prueba Septiembre-2", "Resultados", "Administrativo", "Remuneracion 202609-2 Total.xlsx")),
        ("2026101", Path.Combine("Docs", "Prueba Octubre-1", "Resultados", "Resultado administrativo", "Remuneracion 202610-1 Total.xlsx"))
    ];

    private static string Manual(string periodo) => Ruta(Manuales.Single(m => m.Periodo == periodo).Ruta.Split(Path.DirectorySeparatorChar));

    private static string DocxPlantilla() => Ruta("Docs", "Detalle de plantilla.docx");

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Helper de lectura BCL por bytes (hoja por NOMBRE, nunca por posición)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>Celda de un .xlsx leída por bytes: fórmula (+ shared), tipo, valor y estilo.</summary>
    private sealed record CeldaXlsx(
        string Ref,
        string? Formula,
        string? SiCompartida,
        string? RefCompartida,
        string? Tipo,
        string? Valor);

    private sealed class HojaXlsx
    {
        public HojaXlsx(string nombre, string dimension, IReadOnlyDictionary<string, CeldaXlsx> celdas)
        {
            Nombre = nombre;
            Dimension = dimension;
            Celdas = celdas;
        }

        public string Nombre { get; }

        public string Dimension { get; }

        public IReadOnlyDictionary<string, CeldaXlsx> Celdas { get; }

        public CeldaXlsx? this[string referencia] =>
            Celdas.TryGetValue(referencia, out var celda) ? celda : null;

        /// <summary>Texto de la fórmula efectiva: la propia o la del maestro compartido traducida por columna/fila.</summary>
        public string? FormulaEfectiva(string referencia)
        {
            var celda = this[referencia];
            if (celda is null)
            {
                return null;
            }

            if (!string.IsNullOrEmpty(celda.Formula))
            {
                return celda.Formula;
            }

            if (celda.SiCompartida is null)
            {
                // Sin <f> devuelve null; con <f/> degenerada devuelve "" (no ocurre en v4).
                return celda.Formula;
            }

            // Seguidora de una fórmula compartida: se busca el maestro (mismo si, con texto).
            var maestro = Celdas.Values.FirstOrDefault(c =>
                c.SiCompartida == celda.SiCompartida && !string.IsNullOrEmpty(c.Formula));
            Assert.True(maestro is not null, $"Fórmula compartida si={celda.SiCompartida} sin maestro en '{Nombre}' ({referencia}).");

            var deltaColumnas = ColumnaDe(referencia) - ColumnaDe(maestro!.Ref);
            var deltaFilas = FilaDe(referencia) - FilaDe(maestro.Ref);
            return Trasladar(maestro.Formula!, deltaColumnas, deltaFilas);
        }

        public string? Texto(string referencia) => this[referencia]?.Valor;

        public decimal? Numero(string referencia)
        {
            var valor = this[referencia]?.Valor;
            return decimal.TryParse(valor, NumberStyles.Any, CultureInfo.InvariantCulture, out var numero)
                ? numero
                : null;
        }
    }

    private static HojaXlsx AbrirHoja(string ruta, string nombreHoja)
    {
        using var zip = ZipFile.OpenRead(ruta);

        XDocument workbook;
        using (var stream = zip.GetEntry("xl/workbook.xml")!.Open())
        {
            workbook = XDocument.Load(stream);
        }

        XDocument relaciones;
        using (var stream = zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open())
        {
            relaciones = XDocument.Load(stream);
        }

        var destinos = relaciones.Root!.Elements(Prns + "Relationship")
            .ToDictionary(r => r.Attribute("Id")!.Value, r => r.Attribute("Target")!.Value);

        var hoja = workbook.Root!.Element(Ns + "sheets")!.Elements(Ns + "sheet")
            .SingleOrDefault(s => (string?)s.Attribute("name") == nombreHoja);
        Assert.True(hoja is not null, $"El archivo '{ruta}' no trae la hoja '{nombreHoja}'.");

        var destino = destinos[hoja!.Attribute(Rns + "id")!.Value].Replace("worksheets/", "xl/worksheets/", StringComparison.Ordinal);
        if (!destino.StartsWith("xl/", StringComparison.Ordinal))
        {
            destino = "xl/" + destino;
        }

        XDocument hojaXml;
        using (var stream = zip.GetEntry(destino)!.Open())
        {
            hojaXml = XDocument.Load(stream);
        }

        var compartidas = new List<string>();
        var entradaCompartidas = zip.GetEntry("xl/sharedStrings.xml");
        if (entradaCompartidas is not null)
        {
            XDocument xml;
            using (var stream = entradaCompartidas.Open())
            {
                xml = XDocument.Load(stream);
            }

            foreach (var si in xml.Root!.Elements(Ns + "si"))
            {
                compartidas.Add(string.Concat(si.Descendants(Ns + "t").Select(t => t.Value)));
            }
        }

        var celdas = new Dictionary<string, CeldaXlsx>(StringComparer.Ordinal);
        foreach (var celda in hojaXml.Descendants(Ns + "c"))
        {
            var referencia = (string)celda.Attribute("r")!;
            var formula = celda.Element(Ns + "f");
            var tipo = (string?)celda.Attribute("t");
            var valor = celda.Element(Ns + "v")?.Value;
            if (tipo == "s" && valor is not null)
            {
                valor = compartidas[int.Parse(valor, CultureInfo.InvariantCulture)];
            }

            celdas[referencia] = new CeldaXlsx(
                referencia,
                formula?.Value,
                (string?)formula?.Attribute("si"),
                (string?)formula?.Attribute("ref"),
                tipo,
                valor);
        }

        var dimension = (string?)hojaXml.Descendants(Ns + "dimension").FirstOrDefault()?.Attribute("ref") ?? string.Empty;
        return new HojaXlsx(nombreHoja, dimension, celdas);
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Utilidades de referencia y normalización
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static readonly Regex ReferenciaRegex = new(
        @"(?<![A-Za-z0-9_.$!'])(\$?[A-Z]{1,3}\$?\d+)",
        RegexOptions.CultureInvariant);

    private static string Trasladar(string formula, int deltaColumnas, int deltaFilas)
    {
        if (deltaColumnas == 0 && deltaFilas == 0)
        {
            return formula;
        }

        return ReferenciaRegex.Replace(formula, coincidencia =>
        {
            var token = coincidencia.Value;
            var match = Regex.Match(token, @"^(\$?)([A-Z]{1,3})(\$?)(\d+)$", RegexOptions.CultureInvariant);
            var anclaColumna = match.Groups[1].Value;
            var anclaFila = match.Groups[3].Value;
            var columna = LetraAColumna(match.Groups[2].Value) + (anclaColumna.Length == 0 ? deltaColumnas : 0);
            var fila = int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture) + (anclaFila.Length == 0 ? deltaFilas : 0);
            return anclaColumna + LetraDeColumna(columna) + anclaFila + fila.ToString(CultureInfo.InvariantCulture);
        });
    }

    private static int LetraAColumna(string letra)
    {
        var columna = 0;
        foreach (var caracter in letra)
        {
            columna = (columna * 26) + (caracter - 'A' + 1);
        }

        return columna;
    }

    private static string LetraDeColumna(int columna)
    {
        var letra = string.Empty;
        while (columna > 0)
        {
            var resto = (columna - 1) % 26;
            letra = (char)('A' + resto) + letra;
            columna = (columna - 1) / 26;
        }

        return letra;
    }

    private static int ColumnaDe(string referencia) =>
        LetraAColumna(new string(referencia.Where(char.IsLetter).ToArray()));

    private static int FilaDe(string referencia) =>
        int.Parse(new string(referencia.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);

    /// <summary>Normalización del plan (§2.4): mayúsculas, sin diacríticos, trim y espacios colapsados.</summary>
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
                constructor.Append(caracter);
            }
        }

        return string.Join(' ', constructor.ToString().ToUpperInvariant().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    private static decimal? NumeroDe(object? valor) => valor switch
    {
        null => null,
        double d => (decimal)d,
        float f => (decimal)f,
        decimal m => m,
        int i => i,
        long l => l,
        string s when decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var numero) => numero,
        _ => null
    };

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Lectura de los 30 insumos (ExcelDataReader SOLO sobre COPIA temporal)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private sealed class HojaPlana
    {
        private readonly Dictionary<int, Dictionary<int, object?>> _filas = [];

        public object? this[int fila, string columna]
        {
            get => _filas.TryGetValue(fila, out var celdas) && celdas.TryGetValue(LetraAColumna(columna), out var valor)
                ? valor
                : null;
        }

        public IEnumerable<int> FilasConContenido => _filas.Keys.Order();

        public void Agregar(int fila, int columna, object? valor)
        {
            if (!_filas.TryGetValue(fila, out var celdas))
            {
                celdas = [];
                _filas[fila] = celdas;
            }

            celdas[columna] = valor;
        }
    }

    private static HojaPlana LeerInsumo(string ruta)
    {
        var copia = Path.Combine(Path.GetTempPath(), $"p64-t0-{Guid.NewGuid():N}.xlsx");
        File.Copy(ruta, copia, overwrite: true);
        try
        {
            using var stream = File.Open(copia, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            Assert.Equal("Sheet1", reader.Name);

            var hoja = new HojaPlana();
            var fila = 0;
            while (reader.Read())
            {
                fila++;
                for (var columna = 0; columna < reader.FieldCount; columna++)
                {
                    hoja.Agregar(fila, columna + 1, reader.GetValue(columna));
                }
            }

            Assert.False(reader.NextResult(), "El insumo debe traer una única hoja (Sheet1).");
            return hoja;
        }
        finally
        {
            File.Delete(copia);
        }
    }

    private static string Texto(HojaPlana hoja, int fila, string columna) => Normalizar(hoja[fila, columna] as string);

    private static decimal? Moneda(HojaPlana hoja, int fila, string columna) => NumeroDe(hoja[fila, columna]);

    /// <summary>Marca la fila del marcador por texto normalizado (nunca por posición).</summary>
    private static int FilaDelMarcador(HojaPlana hoja, string marcador)
    {
        var normalizado = Normalizar(marcador);
        var filas = hoja.FilasConContenido
            .Where(f => Enumerable.Range(1, 6).Any(c => Texto(hoja, f, LetraDeColumna(c)) == normalizado))
            .ToArray();
        Assert.True(filas.Length == 1, $"El marcador '{marcador}' aparece {filas.Length} veces (se esperaba exactamente 1).");
        return filas[0];
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Parseo del bloque 2 / bloque 3 de un insumo
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private sealed record FuenteInsumo(
        string Etiqueta,
        decimal Total,
        decimal AplicadoR3,
        decimal AnticiposR3,
        decimal AplicadoR4,
        decimal AnticiposR4,
        IReadOnlyList<string> Canales);

    private sealed record InsumoLeido(
        int Marcador2,
        int Marcador3,
        int FilaEncabezado,
        IReadOnlyList<FuenteInsumo> Fuentes,
        decimal TotalGrande,
        IReadOnlyList<string> ColumnasBloque3,
        decimal TotalBloque3,
        IReadOnlyList<string> FilasNoReconocidas)
    {
        public string FuentesCodificadas =>
            string.Join("|", Fuentes.Select(f => $"{f.Etiqueta}={f.Total.ToString(CultureInfo.InvariantCulture)}"));

        public string Bloque3Codificado => string.Join("|", ColumnasBloque3);
    }

    private static decimal[] AcumuladorNuevo() => [0m, 0m, 0m, 0m];

    private static InsumoLeido ParsearInsumo(HojaPlana hoja)
    {
        var marcador2 = FilaDelMarcador(hoja, MarcadorBloque2);
        var marcador3 = FilaDelMarcador(hoja, MarcadorBloque3);
        Assert.True(marcador2 < marcador3, $"Bloque 3 (r{marcador3}) antes que bloque 2 (r{marcador2}): marcadores invertidos.");

        // Encabezado del bloque 2: fila con E='ASEO' y F='Total' (medido en 30/30).
        var encabezado = Enumerable.Range(marcador2 + 1, marcador3 - marcador2 - 1)
            .Where(f => Texto(hoja, f, "E") == "ASEO" && Texto(hoja, f, "F") == "TOTAL")
            .Single();

        var fuentes = new List<FuenteInsumo>();
        var noReconocidas = new List<string>();
        var acumulador = AcumuladorNuevo();
        var canales = new List<string>();
        var totalGrande = 0m;

        for (var fila = encabezado + 2; fila < marcador3; fila++)
        {
            var a = Texto(hoja, fila, "A");
            var b = Texto(hoja, fila, "B");
            var c = Texto(hoja, fila, "C");
            var d = Moneda(hoja, fila, "D");
            var e = Moneda(hoja, fila, "E");
            var f = Moneda(hoja, fila, "F");

            if (a.Length == 0 && b.Length == 0 && c.Length == 0 && d is null && e is null && f is null)
            {
                continue;
            }

            var indiceConcepto = Array.FindIndex(CatalogoConceptos, x => Normalizar(x.Concepto) == c);
            if (indiceConcepto >= 0)
            {
                Assert.True(a.Length == 0 && b.Length == 0, $"r{fila}: fila de concepto con etiqueta A='{a}' B='{b}'.");
                Assert.True(d is not null && e is not null && f is not null, $"r{fila}: concepto sin las tres columnas D/E/F.");
                Assert.InRange(Math.Abs(d!.Value - e!.Value), 0m, Tolerancia);
                Assert.InRange(Math.Abs(e.Value - f!.Value), 0m, Tolerancia);
                acumulador[indiceConcepto] += d.Value;
                continue;
            }

            if (b.Length > 0 && c == "TOTAL")
            {
                Assert.True(d is not null && e is not null && f is not null, $"r{fila}: total de canal sin D/E/F.");
                Assert.InRange(Math.Abs(d!.Value - e!.Value), 0m, Tolerancia);
                Assert.InRange(Math.Abs(e.Value - f!.Value), 0m, Tolerancia);
                canales.Add(b);
                continue;
            }

            if (a.Length > 0 && b == "TOTAL")
            {
                Assert.True(d is not null && e is not null && f is not null, $"r{fila}: total de fuente sin D/E/F.");
                Assert.InRange(Math.Abs(d!.Value - e!.Value), 0m, Tolerancia);
                Assert.InRange(Math.Abs(e.Value - f!.Value), 0m, Tolerancia);
                fuentes.Add(new FuenteInsumo(a, d.Value, acumulador[0], acumulador[1], acumulador[2], acumulador[3], canales.ToArray()));
                acumulador = AcumuladorNuevo();
                canales.Clear();
                continue;
            }

            if (a == "TOTAL" && b.Length == 0)
            {
                Assert.True(d is not null, $"r{fila}: gran total sin valor D.");
                totalGrande = d!.Value;
                continue;
            }

            noReconocidas.Add($"r{fila}: A='{a}' B='{b}' C='{c}' D='{d}' E='{e}' F='{f}'");
        }

        Assert.Empty(noReconocidas);
        Assert.NotEmpty(fuentes);

        // Bloque 3 (control): encabezados por encabezado, columnas variables por archivo.
        var filaEncabezadoBloque3 = marcador3 + 1;
        var columnas = new List<(int Columna, string Etiqueta)>();
        for (var columna = 1; columna <= 8; columna++)
        {
            var etiqueta = Texto(hoja, filaEncabezadoBloque3, LetraDeColumna(columna));
            if (etiqueta.Length > 0)
            {
                columnas.Add((columna, etiqueta));
            }
        }

        Assert.Contains(columnas, x => x.Etiqueta == "TOTAL");
        var columnaTotal = columnas.Single(x => x.Etiqueta == "TOTAL").Columna;
        var letraTotal = LetraDeColumna(columnaTotal);

        var filaTotalBloque3 = hoja.FilasConContenido
            .Where(f => f > filaEncabezadoBloque3 && Texto(hoja, f, "A") == "TOTAL")
            .DefaultIfEmpty(0)
            .Max();
        Assert.True(filaTotalBloque3 > 0, "El bloque 3 no trae fila TOTAL.");
        var totalBloque3 = Moneda(hoja, filaTotalBloque3, letraTotal) ?? 0m;

        return new InsumoLeido(
            marcador2,
            marcador3,
            encabezado,
            fuentes,
            totalGrande,
            columnas.Select(x => x.Etiqueta).ToArray(),
            totalBloque3,
            noReconocidas);
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Geometría de bloques del template (derivada por ETIQUETAS, nunca por posición)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private sealed record GeometriaBloque(
        int FilaEtiqueta,
        int FilaAses,
        int FilaReportes,
        int FilaAplicado,
        int FilaAnticipos,
        int FilaTotal);

    private static GeometriaBloque GeometriaDe(HojaXlsx hoja, string etiqueta)
    {
        var filaEtiqueta = hoja.Celdas.Values
            .Where(c => ColumnaDe(c.Ref) == 2 && Normalizar(c.Valor) == Normalizar(etiqueta))
            .Select(c => FilaDe(c.Ref))
            .Single();

        var filaAses = filaEtiqueta + 1;
        var filaReportes = filaEtiqueta + 2;
        var filaAplicado = filaEtiqueta + 3;
        var filaAnticipos = filaEtiqueta + 4;
        var filaTotal = filaEtiqueta + 5;

        Assert.Equal("ASE 1", Normalizar(hoja.Texto($"C{filaAses}")));
        Assert.Equal("APLICADO", Normalizar(hoja.Texto($"B{filaAplicado}")));
        Assert.Equal("ANTICIPOS", Normalizar(hoja.Texto($"B{filaAnticipos}")));
        Assert.Equal("TOTAL", Normalizar(hoja.Texto($"B{filaTotal}")));
        Assert.Equal("REPORTE 3", Normalizar(hoja.Texto($"C{filaReportes}")));
        Assert.Equal("REPORTE 4", Normalizar(hoja.Texto($"D{filaReportes}")));

        return new GeometriaBloque(filaEtiqueta, filaAses, filaReportes, filaAplicado, filaAnticipos, filaTotal);
    }

    private static int ColumnaDeAse(HojaXlsx hoja, GeometriaBloque bloque, int aseId)
    {
        var columnas = Enumerable.Range(3, 12)
            .Where(c => Normalizar(hoja.Texto($"{LetraDeColumna(c)}{bloque.FilaAses}")) == $"ASE {aseId}")
            .ToArray();
        Assert.True(columnas.Length == 1, $"ASE {aseId} aparece {columnas.Length} veces en el bloque.");
        var columna = columnas[0];
        Assert.Equal("REPORTE 3", Normalizar(hoja.Texto($"{LetraDeColumna(columna)}{bloque.FilaReportes}")));
        Assert.Equal("REPORTE 4", Normalizar(hoja.Texto($"{LetraDeColumna(columna + 1)}{bloque.FilaReportes}")));
        return columna;
    }

    /// <summary>Referencia de destino del §1.2: etiqueta de bloque + ASE + fila + reporte.</summary>
    private static string Destino(HojaXlsx hoja, string etiquetaBloque, int aseId, string fila, string reporte)
    {
        var bloque = GeometriaDe(hoja, etiquetaBloque);
        var columna = ColumnaDeAse(hoja, bloque, aseId) + (reporte == "REPORTE 4" ? 1 : 0);
        var filaDestino = fila == "APLICADO" ? bloque.FilaAplicado : bloque.FilaAnticipos;
        return $"{LetraDeColumna(columna)}{filaDestino}";
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // T0.1 — Geometría v4 de ANT EXT-REV por bytes (REQ-01)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T0.1a: la hoja existe, su dimensión es <c>A1:P45</c> y las 5 etiquetas de bloque viven en la
    /// columna B en el orden del archivo (B8/B16/B24/B32/B40), cada una con su rejilla
    /// ASE/REPORTE/APLICADO/ANTICIPOS/TOTAL.
    /// </summary>
    [Fact]
    public void T0_01a_HojaAntExtRev_DimensionYBloquesPorEtiqueta()
    {
        var hoja = AbrirHoja(PlantillaV4(), HojaAntExtRev);

        Assert.Equal("A1:P45", hoja.Dimension);

        var etiquetas = hoja.Celdas.Values
            .Where(c => ColumnaDe(c.Ref) == 2)
            .Select(c => (Fila: FilaDe(c.Ref), Texto: Normalizar(c.Valor)))
            .Where(x => EtiquetasBloque.Select(Normalizar).Contains(x.Texto))
            .OrderBy(x => x.Fila)
            .ToArray();

        Assert.Equal(EtiquetasBloque.Select(Normalizar).ToArray(), etiquetas.Select(x => x.Texto).ToArray());
        Assert.Equal([8, 16, 24, 32, 40], etiquetas.Select(x => x.Fila).ToArray());

        foreach (var etiqueta in EtiquetasBloque)
        {
            var bloque = GeometriaDe(hoja, etiqueta);

            // Fila de ASEs: C/E/G/I/K traen ASE 1..ASE 5 y M trae TOTAL.
            Assert.Equal("TOTAL", Normalizar(hoja.Texto($"M{bloque.FilaAses}")));

            // Fila de reportes: alternan REPORTE 3 / REPORTE 4 en C..N.
            for (var columna = 3; columna <= 14; columna++)
            {
                var esperado = columna % 2 == 1 ? "REPORTE 3" : "REPORTE 4";
                Assert.Equal(esperado, Normalizar(hoja.Texto($"{LetraDeColumna(columna)}{bloque.FilaReportes}")));
            }
        }
    }

    /// <summary>
    /// T0.1b: bloque superior (filas 3..6) y controles P3/P4. Cada celda C..L de las filas APLICADO y
    /// ANTICIPOS del bloque superior suma las cinco filas homónimas de los cinco bloques; la fila
    /// TOTAL suma APLICADO+ANTICIPOS y la fila 6 es el auto-chequeo contra la suma de TOTALes.
    /// </summary>
    [Fact]
    public void T0_01b_BloqueSuperior_FormulasYControles()
    {
        var hoja = AbrirHoja(PlantillaV4(), HojaAntExtRev);

        for (var columna = 3; columna <= 12; columna++)
        {
            var col = LetraDeColumna(columna);
            Assert.Equal($"{col}11+{col}19+{col}27+{col}35+{col}43", hoja.FormulaEfectiva($"{col}3"));
            Assert.Equal($"{col}12+{col}20+{col}28+{col}36+{col}44", hoja.FormulaEfectiva($"{col}4"));
            Assert.Equal($"SUM({col}3:{col}4)", hoja.FormulaEfectiva($"{col}5"));
            Assert.Equal($"{col}5=({col}13+{col}21+{col}29+{col}37+{col}45)", hoja.FormulaEfectiva($"{col}6"));
        }

        Assert.Equal("C3+E3+G3+I3+K3", hoja.FormulaEfectiva("M3"));
        Assert.Equal("D3+F3+H3+J3+L3", hoja.FormulaEfectiva("N3"));
        Assert.Equal("C4+E4+G4+I4+K4", hoja.FormulaEfectiva("M4"));
        Assert.Equal("D4+F4+H4+J4+L4", hoja.FormulaEfectiva("N4"));
        Assert.Equal("SUM(M3:M4)", hoja.FormulaEfectiva("M5"));
        Assert.Equal("SUM(N3:N4)", hoja.FormulaEfectiva("N5"));
        Assert.Equal("M5=(M13+M21+M29+M37+M45)", hoja.FormulaEfectiva("M6"));
        Assert.Equal("N5=(N13+N21+N29+N37+N45)", hoja.FormulaEfectiva("N6"));

        // Controles: P3 = (-VALIDACION_TOTAL!M9) y P4 = (-'CONSOLIDADO_TOTAL RECAUDO'!L131).
        Assert.Equal("N3=(-VALIDACION_TOTAL!M9)", hoja.FormulaEfectiva("P3"));
        Assert.Equal("b", hoja["P3"]!.Tipo);
        Assert.Equal("N4=(-'CONSOLIDADO_TOTAL RECAUDO'!L131)", hoja.FormulaEfectiva("P4"));
        Assert.Equal("b", hoja["P4"]!.Tipo);
    }

    /// <summary>
    /// T0.1c: filas APLICADO/ANTICIPOS de cada bloque (rejilla de valores) y fila TOTAL (SUM por
    /// columna). Los M/N de las filas de valor son la suma horizontal de las 5 columnas de ASE; los
    /// M/N del TOTAL congelan la forma heterogénea MEDIDA (bloque 2 y 5 usan suma de subtotales).
    /// </summary>
    [Fact]
    public void T0_01c_RejillaDeValores_FormulasMNYFilaTotal()
    {
        var hoja = AbrirHoja(PlantillaV4(), HojaAntExtRev);

        string[] totalesM =
        [
            "SUM(M11:M12)", "C21+E21+G21+I21+K21", "SUM(M27:M28)", "SUM(M35:M36)", "SUM(M43:M44)"
        ];
        string[] totalesN =
        [
            "SUM(N11:N12)", "SUM(N19:N20)", "SUM(N27:N28)", "SUM(N35:N36)", "SUM(N43:N44)"
        ];

        for (var indice = 0; indice < EtiquetasBloque.Length; indice++)
        {
            var bloque = GeometriaDe(hoja, EtiquetasBloque[indice]);

            foreach (var fila in new[] { bloque.FilaAplicado, bloque.FilaAnticipos })
            {
                Assert.Equal($"C{fila}+E{fila}+G{fila}+I{fila}+K{fila}", hoja.FormulaEfectiva($"M{fila}"));
                Assert.Equal($"D{fila}+F{fila}+H{fila}+J{fila}+L{fila}", hoja.FormulaEfectiva($"N{fila}"));

                for (var columna = 3; columna <= 12; columna++)
                {
                    var col = LetraDeColumna(columna);
                    Assert.True(hoja.FormulaEfectiva($"{col}{fila}") is null, $"{col}{fila} no debe ser fórmula (celda de valor).");
                }
            }

            for (var columna = 3; columna <= 12; columna++)
            {
                var col = LetraDeColumna(columna);
                Assert.Equal(
                    $"SUM({col}{bloque.FilaAplicado}:{col}{bloque.FilaAnticipos})",
                    hoja.FormulaEfectiva($"{col}{bloque.FilaTotal}"));
            }

            Assert.Equal(totalesM[indice], hoja.FormulaEfectiva($"M{bloque.FilaTotal}"));
            Assert.Equal(totalesN[indice], hoja.FormulaEfectiva($"N{bloque.FilaTotal}"));
        }
    }

    /// <summary>
    /// T0.1d: naturaleza de las 100 celdas de valor (5 bloques × 10 columnas C..L × 2 filas
    /// APLICADO/ANTICIPOS). MEDIDO: las 100 son numéricas (<c>&lt;v&gt;0&lt;/v&gt;</c>, sin
    /// <c>t</c> y sin <c>&lt;f&gt;</c>) — no texto ni vacías como decía el §1.1 (hallazgo H0).
    /// </summary>
    [Fact]
    public void T0_01d_CienCeldasDeValor_NaturalezaMedida()
    {
        var hoja = AbrirHoja(PlantillaV4(), HojaAntExtRev);
        var celdas = new List<string>();

        foreach (var etiqueta in EtiquetasBloque)
        {
            var bloque = GeometriaDe(hoja, etiqueta);
            foreach (var fila in new[] { bloque.FilaAplicado, bloque.FilaAnticipos })
            {
                for (var columna = 3; columna <= 12; columna++)
                {
                    var referencia = $"{LetraDeColumna(columna)}{fila}";
                    var celda = hoja[referencia];

                    Assert.True(celda is not null, $"Falta la celda de valor {referencia}.");
                    Assert.Null(celda!.Formula);
                    Assert.Null(celda.SiCompartida);
                    Assert.Null(celda.Tipo);
                    Assert.Equal("0", celda.Valor);
                    celdas.Add(referencia);
                }
            }
        }

        Assert.Equal(100, celdas.Count);
        Assert.Equal(100, celdas.Distinct(StringComparer.Ordinal).Count());
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // T0.2 — ANTICIPOS USUARIOS (100 %-fórmulas) y CONSOLIDADO_TOTAL RECAUDO!L125:L130 (REQ-01)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T0.2a: la rejilla de datos de <c>ANTICIPOS USUARIOS</c> (columnas C/F/I/L/O/R × filas 3..7,
    /// 12..16, 21..25) es fórmula en las 90 celdas y las tres filas de suma son
    /// <c>SUM(Xn:Xn+4)</c> en las 18 celdas (varias seguidoras de fórmulas compartidas, p. ej.
    /// <c>O26</c> = <c>SUM(O21:O25</c>) — el "100 % fórmulas" del §1.1 queda CONFIRMADO.
    /// </summary>
    [Fact]
    public void T0_02a_AnticiposUsuarios_RejillaYSumasTodasDeFormula()
    {
        var hoja = AbrirHoja(PlantillaV4(), HojaAnticipos);

        string[] columnas = ["C", "F", "I", "L", "O", "R"];
        int[] bloques = [3, 12, 21];

        var conFormula = 0;
        var sinFormula = new List<string>();
        foreach (var inicio in bloques)
        {
            for (var fila = inicio; fila < inicio + 5; fila++)
            {
                foreach (var columna in columnas)
                {
                    var referencia = $"{columna}{fila}";
                    if (hoja.FormulaEfectiva(referencia) is null)
                    {
                        sinFormula.Add(referencia);
                    }
                    else
                    {
                        conFormula++;
                    }
                }
            }
        }

        Assert.Equal(90, conFormula);
        Assert.Empty(sinFormula);

        var sumas = 0;
        foreach (var (inicio, filaSuma) in new[] { (3, 8), (12, 17), (21, 26) })
        {
            foreach (var columna in columnas)
            {
                Assert.Equal(
                    $"SUM({columna}{inicio}:{columna}{inicio + 4})",
                    hoja.FormulaEfectiva($"{columna}{filaSuma}"));
                sumas++;
            }
        }

        Assert.Equal(18, sumas);
        Assert.Equal("0", hoja.Texto("O26"));
        Assert.Null(hoja["O26"]!.Tipo);
    }

    /// <summary>
    /// T0.2b: referencias medidas de <c>ANTICIPOS USUARIOS</c> hacia <c>ANT EXT-REV</c> (incluidas las
    /// cuatro citadas por el §1.1: D4/M4/D20/M20) y forma de las sumas y auto-chequeos.
    /// </summary>
    [Fact]
    public void T0_02b_AnticiposUsuarios_ReferenciasAAntExtRev()
    {
        var hoja = AbrirHoja(PlantillaV4(), HojaAnticipos);

        Assert.Equal("'ANT EXT-REV'!C4", hoja.FormulaEfectiva("C12"));
        Assert.Equal("'ANT EXT-REV'!D4", hoja.FormulaEfectiva("C21"));
        Assert.Equal("C17='ANT EXT-REV'!M4", hoja.FormulaEfectiva("C18"));
        Assert.Equal("C26='ANT EXT-REV'!N4", hoja.FormulaEfectiva("C27"));
        Assert.Equal("'ANT EXT-REV'!C20", hoja.FormulaEfectiva("F12"));
        Assert.Equal("'ANT EXT-REV'!D20", hoja.FormulaEfectiva("F21"));
        Assert.Equal("F17='ANT EXT-REV'!M20", hoja.FormulaEfectiva("F18"));
        Assert.Equal("F26='ANT EXT-REV'!N20", hoja.FormulaEfectiva("F27"));
        Assert.Equal("'ANT EXT-REV'!D36", hoja.FormulaEfectiva("I21"));
        Assert.Equal("'ANT EXT-REV'!D44", hoja.FormulaEfectiva("L21"));
        Assert.Equal("'ANT EXT-REV'!C12", hoja.FormulaEfectiva("O12"));
        Assert.Equal("'ANT EXT-REV'!C28", hoja.FormulaEfectiva("R12"));

        // Bloque "anticipos por extemporáneos" (filas 12..16): columna ANTI = fila ANTICIPOS (4) de
        // cada ASE, tomada del reporte REPORTE 3/4 de su par de columnas.
        string[] extemporaneos = ["C4", "E4", "G4", "I4", "K4"];
        for (var indice = 0; indice < 5; indice++)
        {
            Assert.Equal($"'ANT EXT-REV'!{extemporaneos[indice]}", hoja.FormulaEfectiva($"C{12 + indice}"));
        }

        // Bloque "anticipos que cargan por reversiones" (filas 21..25): columna = reporte R4 del ASE.
        string[] cargan = ["D4", "F4", "H4", "J4", "L4"];
        for (var indice = 0; indice < 5; indice++)
        {
            Assert.Equal($"'ANT EXT-REV'!{cargan[indice]}", hoja.FormulaEfectiva($"C{21 + indice}"));
        }

        // Sumas por bloque.
        foreach (var (inicio, filaSuma) in new[] { (3, 8), (12, 17), (21, 26) })
        {
            foreach (var columna in new[] { "C", "F", "I", "L", "O", "R" })
            {
                Assert.Equal($"SUM({columna}{inicio}:{columna}{inicio + 4})", hoja.FormulaEfectiva($"{columna}{filaSuma}"));
            }
        }

        // Auto-chequeos: el primer bloque contra 'REPORTE RECAUDO x BANCO' (su fuente real) y los dos
        // bloques de reversiones contra el M/N del mismo bloque de ANT EXT-REV.
        Assert.Equal("C8='REPORTE RECAUDO x BANCO'!I3", hoja.FormulaEfectiva("C9"));
        Assert.Equal("C17='ANT EXT-REV'!M4", hoja.FormulaEfectiva("C18"));
        Assert.Equal("C26='ANT EXT-REV'!N4", hoja.FormulaEfectiva("C27"));
        Assert.Equal("R17='ANT EXT-REV'!M28", hoja.FormulaEfectiva("R18"));
        Assert.Equal("R26='ANT EXT-REV'!N28", hoja.FormulaEfectiva("R27"));
    }

    /// <summary>
    /// T0.2c: <c>CONSOLIDADO_TOTAL RECAUDO!L125:L130</c> referencia ANTICIPOS USUARIOS. MEDIDO:
    /// L125..L129 son <c>-'ANTICIPOS USUARIOS'!C21..C25</c>; <c>L130</c> está VACÍA (sin fórmula ni
    /// valor) y <c>L131 = SUM(L125:L130)</c> cierra el bloque.
    /// </summary>
    [Fact]
    public void T0_02c_ConsolidadoTotalRecaudo_L125L130_Consumidoras()
    {
        var hoja = AbrirHoja(PlantillaV4(), HojaConsolidadoTotal);

        for (var indice = 0; indice < 5; indice++)
        {
            var fila = 125 + indice;
            Assert.Equal($"'ANTICIPOS USUARIOS'!C{3 + indice}", hoja.FormulaEfectiva($"D{fila}"));
            Assert.Equal($"'ANTICIPOS USUARIOS'!C{12 + indice}", hoja.FormulaEfectiva($"K{fila}"));
            Assert.Equal($"-'ANTICIPOS USUARIOS'!C{21 + indice}", hoja.FormulaEfectiva($"L{fila}"));
        }

        var l130 = hoja["L130"];
        Assert.True(l130 is null || (l130.Formula is null && l130.Valor is null), "L130 debe estar vacía (medido en v4).");
        Assert.Equal("SUM(L125:L130)", hoja.FormulaEfectiva("L131"));
    }

    /// <summary>
    /// T0.2d: la plantilla v4 y las 4 plantillas manuales conservan EXACTAMENTE las mismas 130
    /// fórmulas en <c>ANT EXT-REV</c> y los controles P3/P4 en TRUE (caché del guardado; se mide como
    /// estado del archivo, nunca como oráculo: doctrina Plan 28/29/40). Evidencia de que el manual se
    /// diligenció a valores sin tocar <c>&lt;f&gt;</c>.
    /// </summary>
    [Fact]
    public void T0_02d_Manuales_ConservanLas130FormulasYLosControles()
    {
        var plantilla = AbrirHoja(PlantillaV4(), HojaAntExtRev);
        Assert.Equal(130, plantilla.Celdas.Values.Count(c => c.Formula is not null));

        foreach (var (periodo, _) in Manuales)
        {
            var manual = AbrirHoja(Manual(periodo), HojaAntExtRev);
            Assert.Equal(130, manual.Celdas.Values.Count(c => c.Formula is not null));
            Assert.Equal("1", manual.Texto("P3"));
            Assert.Equal("1", manual.Texto("P4"));
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // T0.3 — Estructura Sheet1 de los 30 insumos + universo de fuentes (REQ-02)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Tabla CONGELADA por H0 (evidencia, nunca contrato de producción): período, ASE, fila de cada
    /// marcador (como evidencia), fuentes del bloque 2 con su total en el orden del archivo, columnas
    /// del bloque 3 y los totales de bloque 2 / bloque 3.
    /// </summary>
    private static readonly (string Periodo, int Ase, int Marcador2, int Marcador3, string Fuentes, string Bloque3, decimal Total2, decimal Total3)[] Estructura30 =
    [
        ("2026071", 1, 11, 33, "OCCIDENTE=-7402760", "RECAUDO OCCIDENTE|TOTAL", -7402760m, -7402760m),
        ("2026071", 2, 54, 104, "ENEL=-9917803|NUEVO ESQUEMA=-620493", "RECAUDO OTROS|TOTAL", -10538296m, -10538296m),
        ("2026071", 3, 52, 104, "ENEL=-5627802|OCCIDENTE=-10051780", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -15679582m, -15679582m),
        ("2026071", 4, 248, 342, "ENEL=-3475400|NUEVO ESQUEMA=0|OCCIDENTE=-522720", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -3998120m, -3998120m),
        ("2026071", 5, 43, 92, "ENEL=-5470988|OCCIDENTE=-1386390", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -6857378m, -6857378m),
        ("2026072", 1, 60, 110, "ENEL=-15501271|OCCIDENTE=-1086060", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -16587331m, -16587331m),
        ("2026072", 2, 65, 124, "ENEL=-9472909|OCCIDENTE=-178170", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -9651079m, -9651079m),
        ("2026072", 3, 47, 94, "ENEL=-30776490|OCCIDENTE=-2650410", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -33426900m, -33426900m),
        ("2026072", 4, 293, 389, "ENEL=-14910059|NUEVO ESQUEMA=0|OCCIDENTE=-14926610", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -29836669m, -29836669m),
        ("2026072", 5, 30, 78, "ENEL=-2114880|OCCIDENTE=-663150", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -2778030m, -2778030m),
        ("2026082", 1, 10, 36, "ENEL=0", "RECAUDO OTROS|TOTAL", 0m, 0m),
        ("2026082", 2, 70, 134, "ENEL=-10736662|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -10736662m, -10736662m),
        ("2026082", 3, 32, 69, "ENEL=-4617354", "RECAUDO OTROS|TOTAL", -4617354m, -4617354m),
        ("2026082", 4, 237, 326, "ENEL=-3442420|NUEVO ESQUEMA=0|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -3442420m, -3442420m),
        ("2026082", 5, 80, 131, "ENEL=-1454550|OCCIDENTE=-66010", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -1520560m, -1520560m),
        ("2026091", 1, 18, 57, "ENEL=-3462950|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -3462950m, -3462950m),
        ("2026091", 2, 30, 72, "ENEL=-8739913", "RECAUDO OTROS|TOTAL", -8739913m, -8739913m),
        ("2026091", 3, 40, 91, "ENEL=-6939290|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -6939290m, -6939290m),
        ("2026091", 4, 346, 433, "ENEL=-3196930|NUEVO ESQUEMA=0|OCCIDENTE=-64000", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -3260930m, -3260930m),
        ("2026091", 5, 31, 72, "ENEL=-8838800|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -8838800m, -8838800m),
        ("2026092", 1, 26, 67, "ENEL=-690790|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -690790m, -690790m),
        ("2026092", 2, 642, 735, "ENEL=-15212817|NUEVO ESQUEMA=0|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -15212817m, -15212817m),
        ("2026092", 3, 40, 84, "ENEL=-6595914", "RECAUDO OTROS|TOTAL", -6595914m, -6595914m),
        ("2026092", 4, 2242, 2358, "ENEL=-4393374|NUEVO ESQUEMA=0|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -4393374m, -4393374m),
        ("2026092", 5, 28, 69, "ENEL=-4790099", "RECAUDO OTROS|TOTAL", -4790099m, -4790099m),
        ("2026101", 1, 30, 79, "ENEL=-3184910|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -3184910m, -3184910m),
        ("2026101", 2, 340, 426, "ENEL=-6297940|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -6297940m, -6297940m),
        ("2026101", 3, 26, 61, "ENEL=-2259272", "RECAUDO OTROS|TOTAL", -2259272m, -2259272m),
        ("2026101", 4, 746, 851, "ENEL=-1581684|NUEVO ESQUEMA=0|OCCIDENTE=0", "RECAUDO OCCIDENTE|RECAUDO OTROS|TOTAL", -1581684m, -1581684m),
        ("2026101", 5, 26, 68, "ENEL=-4005397", "RECAUDO OTROS|TOTAL", -4005397m, -4005397m)
    ];

    /// <summary>
    /// T0.3: los 30 archivos (6 períodos × 5 ASE) traen una única hoja <c>Sheet1</c>, los dos
    /// marcadores en orden, el encabezado del bloque 2 (<c>E='ASEO'</c>, <c>F='Total'</c> y en la fila
    /// siguiente <c>D='66'</c>, <c>E='Total'</c>), filas de concepto / total de canal / total de
    /// fuente con D=E=F, y un bloque 3 con columnas variables cuyo total cierra contra el bloque 2.
    /// </summary>
    [Fact]
    public void T0_03_EstructuraSheet1_DeLos30Insumos()
    {
        Assert.Equal(30, Estructura30.Length);

        var sinOccidente = new List<string>();
        foreach (var esperado in Estructura30)
        {
            var hoja = LeerInsumo(Insumo(esperado.Periodo, esperado.Ase));
            var leido = ParsearInsumo(hoja);

            var contexto = $"{esperado.Periodo} ASE{esperado.Ase}";

            // Marcadores: fila exacta como EVIDENCIA (no contrato) y orden bloque 2 → bloque 3.
            Assert.Equal(esperado.Marcador2, leido.Marcador2);
            Assert.Equal(esperado.Marcador3, leido.Marcador3);
            Assert.True(leido.Marcador2 < leido.Marcador3, $"{contexto}: marcadores invertidos.");

            // Encabezado del bloque 2 (HUE-3): E='ASEO', F='Total'; H+1: D='66', E='Total'.
            Assert.Equal("ASEO", Texto(hoja, leido.FilaEncabezado, "E"));
            Assert.Equal("TOTAL", Texto(hoja, leido.FilaEncabezado, "F"));
            Assert.Equal("66", Texto(hoja, leido.FilaEncabezado + 1, "D"));
            Assert.Equal("TOTAL", Texto(hoja, leido.FilaEncabezado + 1, "E"));

            // Universo de fuentes por archivo, en el orden del archivo.
            Assert.Equal(esperado.Fuentes, leido.FuentesCodificadas);
            Assert.NotEmpty(leido.Fuentes);

            // Bloque 3 (control): columnas variables, nunca fuente.
            Assert.Equal(esperado.Bloque3, leido.Bloque3Codificado);

            // Totales: el bloque 3 cierra contra el bloque 2 del archivo.
            Assert.Equal(esperado.Total2, leido.TotalGrande);
            Assert.Equal(esperado.Total3, leido.TotalBloque3);
            Assert.InRange(leido.TotalBloque3 - leido.TotalGrande, -Tolerancia, Tolerancia);

            // Invariante Σ-grupo == total de fuente (ya validado dentro de ParsearInsumo por D/E/F;
            // aquí se re-expresa sobre la tabla congelada).
            Assert.Equal(esperado.Total2, leido.Fuentes.Sum(f => f.Total));

            if (!leido.Bloque3Codificado.Contains("RECAUDO OCCIDENTE", StringComparison.Ordinal))
            {
                sinOccidente.Add(contexto);
            }
        }

        // MEDIDO: 8 de los 30 archivos no traen columna OCCIDENTE en el bloque 3 (no solo ASE2
        // Sept-1 como citaba el §1.3): se congela la lista completa.
        Assert.Equal(
            ["2026071 ASE2", "2026082 ASE1", "2026082 ASE3", "2026091 ASE2", "2026092 ASE3", "2026092 ASE5", "2026101 ASE3", "2026101 ASE5"],
            sinOccidente.ToArray());
    }

    /// <summary>
    /// T0.3b: la aritmética interna del bloque 2 cierra en los 30 archivos — por cada fuente
    /// presente, Σ de sus 4 conceptos == total de fuente ±0,5 (invariante §1.4) y cada total de canal
    /// es la suma de los conceptos que lo preceden ±0,5.
    /// </summary>
    [Fact]
    public void T0_03b_InvarianteSigmaConceptos_IgualTotalFuente()
    {
        foreach (var (periodo, _) in Periodos)
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                var leido = ParsearInsumo(LeerInsumo(Insumo(periodo, ase)));
                foreach (var fuente in leido.Fuentes)
                {
                    var suma = fuente.AplicadoR3 + fuente.AnticiposR3 + fuente.AplicadoR4 + fuente.AnticiposR4;
                    Assert.InRange(
                        suma - fuente.Total,
                        -Tolerancia,
                        Tolerancia);
                    Assert.NotEmpty(fuente.Canales);
                }

                Assert.Equal(leido.TotalGrande, leido.Fuentes.Sum(f => f.Total));
            }
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // T0.4 — Universo de fuentes 30/30 + ausencia de CIUDAD LIMPIA / ENERBIT (REQ-02 + HUE-1)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T0.4: en los 30 insumos solo aparecen ENEL, NUEVO ESQUEMA y OCCIDENTE como etiquetas de fuente
    /// del bloque 2; EAAB CIUDAD LIMPIA y ENERBIT no aparecen nunca —ni en el bloque 2 ni en el
    /// bloque 3— y están en 0 en las 4 plantillas manuales (bloques B24 y B40, 5 ASE cada una).
    /// </summary>
    [Fact]
    public void T0_04_UniversoDeFuentes_SoloTresEtiquetas()
    {
        var universo = new SortedSet<string>(StringComparer.Ordinal);
        var conteo = new Dictionary<string, int>(StringComparer.Ordinal);

        foreach (var esperado in Estructura30)
        {
            var leido = ParsearInsumo(LeerInsumo(Insumo(esperado.Periodo, esperado.Ase)));
            foreach (var fuente in leido.Fuentes)
            {
                universo.Add(fuente.Etiqueta);
                conteo[fuente.Etiqueta] = conteo.GetValueOrDefault(fuente.Etiqueta) + 1;
                foreach (var ausente in FuentesDeclaradasAusentes)
                {
                    Assert.NotEqual(Normalizar(ausente), fuente.Etiqueta);
                }
            }

            foreach (var ausente in FuentesDeclaradasAusentes)
            {
                Assert.DoesNotContain(Normalizar(ausente), leido.Bloque3Codificado, StringComparison.Ordinal);
            }
        }

        Assert.Equal(["ENEL", "NUEVO ESQUEMA", "OCCIDENTE"], universo.ToArray());
        Assert.Equal(29, conteo["ENEL"]);
        Assert.Equal(22, conteo["OCCIDENTE"]);
        Assert.Equal(8, conteo["NUEVO ESQUEMA"]);

        // 0 en los 4 manuales: bloque EAAB CIUDAD LIMPIA (B24) y bloque ENERBIT (B40).
        foreach (var (periodo, _) in Manuales)
        {
            var manual = AbrirHoja(Manual(periodo), HojaAntExtRev);
            foreach (var ausente in FuentesDeclaradasAusentes)
            {
                var bloque = GeometriaDe(manual, ausente);
                for (var ase = 1; ase <= 5; ase++)
                {
                    var columna = ColumnaDeAse(manual, bloque, ase);
                    foreach (var fila in new[] { bloque.FilaAplicado, bloque.FilaAnticipos })
                    {
                        foreach (var desplazamiento in new[] { 0, 1 })
                        {
                            var referencia = $"{LetraDeColumna(columna + desplazamiento)}{fila}";
                            Assert.True(
                                (manual.Numero(referencia) ?? 0m) == 0m,
                                $"{periodo}: el bloque '{ausente}' debe estar en 0 ({referencia}).");
                        }
                    }
                }
            }
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // T0.5 — Paridad numérica del §1.4 re-derivada (REQ-03)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T0.5a: los cuatro casos citados por el §1.4 se re-derivan del insumo y se comparan con el
    /// manual del administrativo, celda por celda. MEDIDO: coinciden todos (Δ &lt; 1e-6).
    /// </summary>
    [Fact]
    public void T0_05a_ParidadCitadaPorElPlan_Sept1_Julio2_Octubre1()
    {
        var casos = new (string Periodo, int Ase, string Fuente, string EtiquetaBloque, decimal AplicadoR3, decimal AnticiposR3, decimal AplicadoR4, decimal AnticiposR4)[]
        {
            ("2026091", 1, "OCCIDENTE", "BANCO DE OCCIDENTE", 42680m, 40630m, -83310m, 0m),
            ("2026091", 1, "ENEL", "ENEL", 0m, 0m, -3461197.29m, -1752.71m),
            ("2026091", 4, "OCCIDENTE", "BANCO DE OCCIDENTE", 239834m, 1438701m, -1742535m, 0m),
            ("2026091", 4, "NUEVO ESQUEMA", "EAAB RECIPROCIDAD", 208970m, 53030m, -262000m, 0m),
            ("2026091", 4, "ENEL", "ENEL", 8186910m, 9321147m, -20245524.78m, -459462.22m)
        };

        foreach (var caso in casos)
        {
            var leido = ParsearInsumo(LeerInsumo(Insumo(caso.Periodo, caso.Ase)));
            var fuente = leido.Fuentes.Single(f => f.Etiqueta == caso.Fuente);

            Assert.InRange(fuente.AplicadoR3 - caso.AplicadoR3, -Tolerancia, Tolerancia);
            Assert.InRange(fuente.AnticiposR3 - caso.AnticiposR3, -Tolerancia, Tolerancia);
            Assert.InRange(fuente.AplicadoR4 - caso.AplicadoR4, -Tolerancia, Tolerancia);
            Assert.InRange(fuente.AnticiposR4 - caso.AnticiposR4, -Tolerancia, Tolerancia);

            var manual = AbrirHoja(Manual(caso.Periodo), HojaAntExtRev);
            Assert.InRange(
                (manual.Numero(Destino(manual, caso.EtiquetaBloque, caso.Ase, "APLICADO", "REPORTE 3")) ?? 0m) - caso.AplicadoR3,
                -Tolerancia,
                Tolerancia);
        }
    }

    /// <summary>
    /// T0.5b: paridad completa <c>fuente × etiqueta</c> en los 4 períodos con manual (20 archivos):
    /// las 4 celdas del §1.2 por fuente presente coinciden con el manual ±0,5. La única celda vacía
    /// del universo de manuales es <c>2026092!J20</c> (concepto ausente que el manual deja en blanco
    /// en lugar de 0 explícito) — se declara y se compara como 0.
    /// </summary>
    [Theory]
    [InlineData("2026072")]
    [InlineData("2026091")]
    [InlineData("2026092")]
    [InlineData("2026101")]
    public void T0_05b_ParidadCompleta_FuentePorEtiqueta_ContraElManual(string periodo)
    {
        var manual = AbrirHoja(Manual(periodo), HojaAntExtRev);
        var vacias = new List<string>();

        for (var ase = 1; ase <= 5; ase++)
        {
            var leido = ParsearInsumo(LeerInsumo(Insumo(periodo, ase)));
            foreach (var fuente in leido.Fuentes)
            {
                var par = CatalogoBloqueFuente.Single(x => x.Fuente == fuente.Etiqueta);

                var celdas = new (string Fila, string Reporte, decimal Valor)[]
                {
                    ("APLICADO", "REPORTE 3", fuente.AplicadoR3),
                    ("ANTICIPOS", "REPORTE 3", fuente.AnticiposR3),
                    ("APLICADO", "REPORTE 4", fuente.AplicadoR4),
                    ("ANTICIPOS", "REPORTE 4", fuente.AnticiposR4)
                };

                foreach (var (fila, reporte, valor) in celdas)
                {
                    var destino = Destino(manual, par.EtiquetaBloque, ase, fila, reporte);
                    var celda = manual[destino];
                    if (celda?.Valor is null)
                    {
                        vacias.Add($"{periodo}!{destino}");
                        Assert.Equal(0m, valor);
                    }

                    var enManual = manual.Numero(destino) ?? 0m;
                    Assert.InRange(enManual - valor, -Tolerancia, Tolerancia);
                }
            }
        }

        var esperadas = periodo == "2026092" ? new[] { "2026092!J20" } : Array.Empty<string>();
        Assert.Equal(esperadas, vacias.ToArray());
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // T0.6 — Divergencia docx-vs-encargo (D-10) con cita exacta
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T0.6: el docx RECTOR dice literalmente que la hoja se alimenta del "Resumen Recaudo Aplicado
    /// Por Servicio" del reporte de recaudos reversados y que se entregará DILIGENCIADA; el bloque 3
    /// medido NO discrimina por fuente (columnas RECAUDO OCCIDENTE / RECAUDO OTROS / Total, con
    /// archivos sin OCCIDENTE) y el bloque 2 sí trae una fila de total por fuente ⇒ el veredicto
    /// D-10 (bloque 2 = fuente, bloque 3 = control) queda sustentado por bytes.
    /// </summary>
    [Fact]
    public void T0_06_DivergenciaDocxVsEncargo_Bloque2Fuente()
    {
        var texto = TextoDocx(DocxPlantilla());

        Assert.Contains("ANT EXT-REV", texto, StringComparison.Ordinal);
        Assert.Contains(
            "Se utiliza la información \"Resumen Recaudo Aplicado Por Servicio \", la cual se encuentra al final del reporte de recaudos reversados generado para cada concesionario o ASE.",
            texto,
            StringComparison.Ordinal);
        Assert.Contains(
            "La información de esta hoja \"ANT EXT-REV\" se utiliza o alimenta la hoja de \"ANTICIPOS USUARIOS\".",
            texto,
            StringComparison.Ordinal);
        Assert.Contains("hoja \"ANT EXT-REV\" diligenciada", texto, StringComparison.Ordinal);

        // El bloque 3 (lo que el docx nombra como fuente) no puede discriminar por fuente: sus
        // encabezados son RECAUDO OCCIDENTE / RECAUDO OTROS / Total, nunca ENEL/NUEVO ESQUEMA/ENERBIT.
        foreach (var esperado in Estructura30)
        {
            Assert.DoesNotContain("ENEL", esperado.Bloque3, StringComparison.Ordinal);
            Assert.DoesNotContain("NUEVO ESQUEMA", esperado.Bloque3, StringComparison.Ordinal);
            Assert.DoesNotContain("ENERBIT", esperado.Bloque3, StringComparison.Ordinal);

            // El bloque 2 sí: cada archivo trae filas de total rotuladas por fuente del catálogo.
            foreach (var fuente in esperado.Fuentes.Split('|'))
            {
                var etiqueta = fuente.Split('=')[0];
                Assert.Contains(etiqueta, new[] { "ENEL", "NUEVO ESQUEMA", "OCCIDENTE" });
            }
        }
    }

    /// <summary>Texto plano del docx (párrafos separados por salto de línea, comillas tipográficas normalizadas).</summary>
    private static string TextoDocx(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        XDocument documento;
        using (var stream = zip.GetEntry("word/document.xml")!.Open())
        {
            documento = XDocument.Load(stream);
        }

        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var parrafos = documento.Descendants(w + "p")
            .Select(p => string.Concat(p.Descendants(w + "t").Select(t => t.Value)))
            .Where(p => p.Length > 0)
            .ToArray();

        return NormalizarParrafo(string.Join("\n", parrafos));
    }

    private static string NormalizarParrafo(string texto) => texto
        .Replace('\u201C', '"')
        .Replace('\u201D', '"')
        .Replace('\u2019', '\'')
        .Replace('\u00A0', ' ');

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // T0.7 — Huella v4 antes/después (S0 G-HUELLA)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T0.7: la huella semántica de la plantilla v4 es la citada por el plan
    /// (<c>a7476ee2…8c55</c>). H0 no escribe la plantilla: este assert es el "antes" y el "después".
    /// </summary>
    [Fact]
    public void T0_07_HuellaV4_Intacta()
    {
        var huella = HuellaSemanticaPlantilla.Calcular(PlantillaV4());

        Assert.Equal("a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55", huella);
        Assert.Equal(HuellaSemanticaPlantilla.V4, huella);

        // El SHA del archivo es informativo (no gate): se reporta para la evidencia.
        Assert.Equal(64, HuellaSemanticaPlantilla.Sha256Archivo(PlantillaV4()).Length);
    }
}


