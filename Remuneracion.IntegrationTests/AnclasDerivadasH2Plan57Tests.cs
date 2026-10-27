using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 57 (H2.1-H2.3, S3.1-S3.2): cobertura de las anclas EXTEMP <b>derivadas</b> de la plantilla y
/// candado anti-quemado de dos capas.
///
/// <list type="bullet">
/// <item><b>15 casos (5 bloques × 3 períodos):</b> el ancla del bloque es la celda que declara el
/// <c>&lt;f&gt;</c> de la publicación del CONSOLIDADO, su naturaleza se lee del workbook y la
/// descomposición sale de su propio <c>&lt;f&gt;</c>.</item>
/// <item><b>Vía literal:</b> un ancla sin <c>&lt;f&gt;</c> se LEE (la guarda «trae fórmula =
/// ERR-PLANTILLA» del Plan 49 ya no existe).</item>
/// <item><b>Irresoluble:</b> una referencia a una celda inexistente produce error accionable
/// (ASE + hoja + concepto) y cero escritura.</item>
/// <item><b>Candado estático:</b> el camino del oráculo no contiene anclas por celda/ASE, literales de
/// celda-ancla, nombres de empresa ni ramas por quincena; allow-list cerrada + injerto que enrojece.</item>
/// <item><b>Candado conductual:</b> una fórmula injertada con un término falso dispara el núcleo con
/// el Δ exacto, por la misma ruta y sin tabla.</item>
/// <item><b>CA-A6:</b> el guard anti-fórmula del ESCRITOR sigue abortando (no se relajó nada).</item>
/// </list>
///
/// <para>
/// Los tres períodos reales se procesan UNA vez cada uno (caché por período): la corrida de 2026071 y
/// 2026101 termina con la divergencia de espejo PRE-EXISTENTE (<c>O25/D34/F34</c>, fuera del alcance
/// de este plan) en la etapa de validación —el workbook marcado ya está escrito—, así que el fixture
/// lee el oráculo de ese archivo, igual que la corrida real.
/// </para>
/// </summary>
[Collection("Cli")]
public sealed class AnclasDerivadasH2Plan57Tests
{
    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    private const string HojaR1 = "Reporte Componentes R1";

    private static readonly object Candado = new();
    private static readonly Dictionary<string, Corrida> Corridas = new(StringComparer.Ordinal);

    private sealed record Corrida(
        string Ruta,
        IReadOnlyList<ValidacionCruzadaSnapshot> Snapshots,
        ResultadoRemuneracion? Resultado,
        IReadOnlyList<WorkbookLeafInputs>? Leafs);

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S3.1 — 5 bloques × 3 períodos: ancla derivada + naturaleza leída
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Theory]
    [InlineData("2026071", 1)]
    [InlineData("2026071", 2)]
    [InlineData("2026071", 3)]
    [InlineData("2026071", 4)]
    [InlineData("2026071", 5)]
    [InlineData("2026072", 1)]
    [InlineData("2026072", 2)]
    [InlineData("2026072", 3)]
    [InlineData("2026072", 4)]
    [InlineData("2026072", 5)]
    [InlineData("2026101", 1)]
    [InlineData("2026101", 2)]
    [InlineData("2026101", 3)]
    [InlineData("2026101", 4)]
    [InlineData("2026101", 5)]
    public void H2_1_AnclaDelBloque_EsLaReferenciaPublicadaYNaturalezaLeida(string periodo, int aseId)
    {
        var corrida = Correr(periodo);
        var evidencias = corrida.Snapshots.Single(s => s.Ase.Id == aseId).EvidenciaAnclasExtemp;

        // v2: el bloque publica su EXTEMP en el consolidado y en las cinco hojas publicadoras.
        Assert.Equal(6, evidencias.Count);

        // El PRIMER destino es el que declara la celda de publicación del bloque (posición LEÍDA).
        var publicada = ReferenciaPublicada(corrida.Ruta, aseId);
        Assert.Equal(publicada, (evidencias[0].Hoja, evidencias[0].Celda));

        // Naturaleza en runtime: fórmula → se evalúa con los valores del workbook; literal → se lee.
        Assert.False(string.IsNullOrWhiteSpace(evidencias[0].ExpresionSecciones));
        foreach (var evidencia in evidencias)
        {
            if (string.IsNullOrWhiteSpace(evidencia.ExpresionSecciones))
            {
                Assert.Null(evidencia.ValorSecciones);
                continue;
            }

            Assert.NotNull(evidencia.ValorSecciones);
            Assert.Equal(evidencia.ValorAncla, evidencia.ValorSecciones!.Value);
        }

        // El núcleo sale de la MISMA referencia publicada (D{46+aseId}) y también es resoluble.
        Assert.NotNull(corrida.Snapshots.Single(s => s.Ase.Id == aseId).ExtempResuelto);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S3.1 — vía literal (la guarda «trae fórmula» del Plan 49 ya no existe)
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H2_1_AnclaLiteral_SeLee_SinGuardaDeFormula()
    {
        var corrida = Correr("2026072");
        var periodo = Periodo.Parse("2026072");
        using var copia = new CopiaTemporal(corrida.Ruta);
        var (hoja, celda) = ReferenciaPublicada(copia.Ruta, 5);

        const decimal valorLiteral = 12345.5m;
        ConvertirEnLiteral(copia.Ruta, hoja, celda, valorLiteral);

        var snapshot = new ValidacionOracleReader().LeerSnapshots(copia.Ruta, periodo).Single(s => s.Ase.Id == 5);
        var ancla = snapshot.EvidenciaAnclasExtemp.First(a => a.Celda == celda && a.Hoja == hoja);

        Assert.Equal(string.Empty, ancla.ExpresionSecciones);
        Assert.Null(ancla.ValorSecciones);
        Assert.Equal(valorLiteral, ancla.ValorAncla);
        Assert.Equal(valorLiteral, snapshot.ExtempResuelto);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S3.1 — irresoluble: error accionable (concepto + bloque + hoja) y cero escritura
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H2_1_AnclaIrresoluble_ErrorAccionableYCeroEscritura()
    {
        var corrida = Correr("2026072");
        var periodo = Periodo.Parse("2026072");
        using var copia = new CopiaTemporal(corrida.Ruta);

        // La publicación del bloque 5 apunta a una celda que no existe en la hoja de componentes.
        EscribirFormula(copia.Ruta, HojaConsolidado, "D51", $"'Reporte Componentes R1'!ZZ999");
        var shaAntes = Sha256(copia.Ruta);

        var ex = Assert.Throws<CalculoInvalidoException>(
            () => new ValidacionOracleReader().LeerSnapshots(copia.Ruta, periodo));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("ASE 5", ex.Message, StringComparison.Ordinal);
        Assert.Contains(HojaR1, ex.Message, StringComparison.Ordinal);
        Assert.Contains("EXTEMPORANEO", ex.Message, StringComparison.Ordinal);
        Assert.Contains("D51", ex.Message, StringComparison.Ordinal);

        // Mensaje del estado del workbook, no historia del código (Plan 57 §2-D5).
        Assert.DoesNotContain("congel", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("T0", ex.Message, StringComparison.Ordinal);

        // El oráculo es read-only: el workbook no cambia por leerlo.
        Assert.Equal(shaAntes, Sha256(copia.Ruta));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S3.2 (CA-A6) — el guard anti-fórmula del ESCRITOR sigue intacto
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H2_2_CaA6_EscritorSobreCeldaConFormula_SigueAbortando()
    {
        var corrida = Correr("2026072");
        using var copia = new CopiaTemporal(corrida.Ruta);

        // Destino del slot F8 con <f> (mismo escenario que el guard del escritor protege).
        EscribirFormula(copia.Ruta, HojaR1, "F8", "1+1");
        var formulasAntes = FotoFormulas(copia.Ruta, HojaR1);

        var bloque = new BloqueEspejoAseInputs
        {
            Ase = AseFactory.DesdeId(1),
            TieneColumnaEspeciales = false,
            Filas =
            [
                new FilaEspejoR1
                {
                    C = "ENEL",
                    D = "Total",
                    ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase) { ["Total"] = 9m }
                }
            ]
        };

        using (var documento = SpreadsheetDocument.Open(copia.Ruta, true))
        {
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var ex = Assert.Throws<CalculoInvalidoException>(() => EscritorValoresSlotR1.Escribir(parte, bloque));
            Assert.Equal(CodigoError.Plantilla, ex.Codigo);
            Assert.Contains("F8", ex.Message, StringComparison.Ordinal);
        }

        // Cero escritura parcial: las <f> quedan byte-idénticas.
        Assert.Equal(formulasAntes, FotoFormulas(copia.Ruta, HojaR1));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // D6 conductual — injerto de un término falso: el núcleo dispara con el Δ exacto
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H2_1_ConductualFormulaInjertada_DisparaNucleoConDeltaExacto()
    {
        // Plan 59 (§11-bis): se corre sobre 2026101. El conductor (injerto de un término falso → el
        // núcleo dispara con el Δ exacto) necesita una corrida cuyo resultado exista y no depende del
        // período; se conserva el período con el que se publicó la evidencia. (El motivo original era
        // que 2026072 no cerraba por el typo de plantilla R1!F328 —corregido en la v4—; mover la corrida
        // otra vez solo cambiaría la evidencia, no la cobertura.)
        var corrida = Correr("2026101");
        Assert.NotNull(corrida.Resultado);
        Assert.NotNull(corrida.Leafs);

        var periodo = Periodo.Parse("2026101");
        using var copia = new CopiaTemporal(corrida.Ruta);
        var (hoja, celda) = ReferenciaPublicada(copia.Ruta, 5);

        const decimal terminoFalso = 1000m;
        AnexarTermino(copia.Ruta, hoja, celda, terminoFalso);

        var snapshots = new ValidacionOracleReader().LeerSnapshots(copia.Ruta, periodo);
        var snapshotAse5 = snapshots.Single(s => s.Ase.Id == 5);

        // La evidencia del interior sigue siendo consistente (es diagnóstico, §2-D4.2): el ancla
        // evalúa su propia fórmula. Lo que cambia es el valor visible del bloque.
        var ancla = snapshotAse5.EvidenciaAnclasExtemp.First(a => a.Celda == celda && a.Hoja == hoja);
        Assert.Equal(ancla.ValorAncla, ancla.ValorSecciones);

        var errores = new ValidadorBasico().Validar(corrida.Resultado!, corrida.Leafs!, snapshots);

        var delNucleo = errores
            .Where(e => e.Contains($"{HojaConsolidado} · D51", StringComparison.Ordinal))
            .ToList();
        var errorNucleo = Assert.Single(delNucleo);
        Assert.Contains("ASE 5", errorNucleo, StringComparison.Ordinal);
        Assert.Contains("EXTEMP resuelto desde lo escrito", errorNucleo, StringComparison.Ordinal);
        Assert.Contains($"Diferencia={terminoFalso}", errorNucleo, StringComparison.Ordinal);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // D6 estático — candado anti-quemado + prueba de injerto
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H2_1_CandadoEstatico_CaminoDelOraculoSinAnclasQuemadas()
    {
        foreach (var (archivo, desde, hasta) in CaminoDelOraculo)
        {
            var fuente = desde is null ? File.ReadAllText(Ruta(archivo)) : Region(Ruta(archivo), desde, hasta!);
            var violaciones = Analizar(fuente);

            Assert.True(
                violaciones.Count == 0,
                $"{archivo} [{desde ?? "completo"}] viola G-D-H/CA-A4 (anclas quemadas en el camino del oráculo):"
                + Environment.NewLine + string.Join(Environment.NewLine, violaciones));
        }
    }

    [Fact]
    public void H2_1_ElMapaCongeladoDeAnclas_NoExisteEnProduccion()
    {
        var raiz = RaizRepo();
        var produccion = new[]
        {
            Path.Combine(raiz, "Remuneracion.Core"),
            Path.Combine(raiz, "Remuneracion.Infrastructure"),
            Path.Combine(raiz, "Remuneracion.WinForms"),
            Path.Combine(raiz, "Remuneracion.Cli")
        };

        var referencias = produccion
            .SelectMany(dir => Directory.EnumerateFiles(dir, "*.cs", SearchOption.AllDirectories))
            .Where(archivo => !archivo.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.Ordinal))
            .Where(archivo => !archivo.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.Ordinal))
            .Select(archivo => (Archivo: archivo, Codigo: string.Join('\n', SinComentarios(File.ReadAllText(archivo)))))
            .Where(par => Regex.IsMatch(par.Codigo, @"\bAnclasExtempR1\b"))
            .Select(par => Path.GetRelativePath(raiz, par.Archivo))
            .ToList();

        // Plan 57 (§2-D2.3): la tabla congelada se RETIRÓ; `AnclaExtempR1Evidencia` es el snapshot de
        // evidencia derivada y no es la tabla (el filtro excluye `Ancla` sin `s`+`ExtempR1`).
        Assert.Empty(referencias);
    }

    [Fact]
    public void H2_1_Injerto_AnclaPorCeldaOEmpresa_Enrojece()
    {
        Assert.NotEmpty(Analizar("""var anclas = new Dictionary<int, string> { [5] = "F682" };"""));
        Assert.NotEmpty(Analizar("""var ancla = "F682";"""));
        Assert.NotEmpty(Analizar("""if (aseId == 5) { celda = "F702"; }"""));
        Assert.NotEmpty(Analizar("""var etiqueta = "OCCIDENTE";"""));
        Assert.NotEmpty(Analizar("""if (periodo.NumeroQuincena == 2) { return; }"""));
        Assert.NotEmpty(Analizar("""throw new InvalidOperationException("el contrato T0 la congeló como literal");"""));
    }

    [Fact]
    public void H2_1_ElAnalizador_NoEnrojeceConLaMaquinariaPermitida()
    {
        // Allow-list cerrada: seguir la referencia declarada, evaluar/leer del workbook y nombrar la
        // hoja de publicación. Todo esto ES la resolución derivada (§2-D6.1).
        const string permitido = """
            var texto = resolutor.TextoFormulaEfectiva(PublicacionExtempConsolidado.HojaConsolidado, celda);
            var ok = resolutor.TryReferenciaUnica(nombre, texto, out var hoja, out var celda);
            var celdaPublicacion = $"D{PublicacionExtempConsolidado.FilaConsolidadoExtemp(aseId)}";
            if (!resolutor.TryEvaluar(hoja, textoFormula, out var secciones)) { return []; }
            foreach (var nombre in resolutor.NombresHojas()) { }
            """;

        Assert.Empty(Analizar(permitido));
    }

    // ── candado: archivos y regiones del camino del oráculo ────────────────────────────────────

    private static readonly (string Archivo, string? Desde, string? Hasta)[] CaminoDelOraculo =
    [
        (Path.Combine("Remuneracion.Core", "Models", "PublicacionExtempConsolidado.cs"), null, null),
        (
            Path.Combine("Remuneracion.Infrastructure", "Excel", "ValidacionOracleReader.cs"),
            "private static decimal ResolverExtempConsolidado(",
            "private static (bool Computable, decimal Izquierdo, decimal Derecho) ResolverOperandos("
        ),
        (
            Path.Combine("Remuneracion.Core", "Services", "ValidadorBasico.cs"),
            "private static void ValidarExtempAbsolutoPorAse(",
            "private static void ValidarBalanceSc("
        )
    ];

    /// <summary>
    /// Analiza el CÓDIGO (sin comentarios) del camino del oráculo y devuelve las violaciones: anclas
    /// por celda/ASE, literales de celda-ancla, nombres de empresa, ramas por quincena y el mensaje
    /// retirado.
    /// </summary>
    internal static IReadOnlyList<string> Analizar(string fuente)
    {
        var violaciones = new List<string>();
        var lineas = SinComentarios(fuente);

        for (var i = 0; i < lineas.Count; i++)
        {
            var linea = lineas[i];

            if (Regex.IsMatch(linea, "\"[A-Z]{1,3}\\d+\""))
            {
                violaciones.Add($"línea {i + 1}: literal de celda-ancla → {linea.Trim()}");
            }

            if (Regex.IsMatch(linea, @"\bAnclasExtempR1\b"))
            {
                violaciones.Add($"línea {i + 1}: consume la tabla congelada de anclas → {linea.Trim()}");
            }

            if (Regex.IsMatch(linea, @"\bNumeroQuincena\b|\bQuincena\b"))
            {
                violaciones.Add($"línea {i + 1}: rama por quincena → {linea.Trim()}");
            }

            if (Regex.IsMatch(linea, @"\baseId\s*switch\b|\bcase\s+[1-5]\s*:|\baseId\s*==\s*[1-5]\b"))
            {
                violaciones.Add($"línea {i + 1}: rama/lista por ASE → {linea.Trim()}");
            }

            if (linea.Contains("congel", StringComparison.OrdinalIgnoreCase))
            {
                violaciones.Add($"línea {i + 1}: mensaje retirado (la foto congelada) → {linea.Trim()}");
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

    private static readonly string[] RotulosDeEmpresa =
    [
        "ENEL", "OCCIDENTE", "ENERBIT", "RECIPROCIDAD", "CIUDAD LIMPIA-ACUEDUCTO", "EAB"
    ];

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

    // ════════════════════════════════════════════════════════════════════════════════════════
    // Corridas reales (una por período, cacheadas) y utilidades de workbook
    // ════════════════════════════════════════════════════════════════════════════════════════

    private static Corrida Correr(string codigoPeriodo)
    {
        lock (Candado)
        {
            if (Corridas.TryGetValue(codigoPeriodo, out var existente))
            {
                return existente;
            }

            var periodo = Periodo.Parse(codigoPeriodo);
            var directorio = Path.Combine(
                Path.GetTempPath(), $"p57-h2-{codigoPeriodo}-{Environment.ProcessId}-{Guid.NewGuid():N}");
            Directory.CreateDirectory(directorio);
            var salida = Path.Combine(directorio, periodo.NombreArchivo);

            ResultadoRemuneracion? resultado = null;
            IReadOnlyList<WorkbookLeafInputs>? leafs = null;
            try
            {
                var proceso = CrearProcesador().Ejecutar(new SolicitudProcesoPeriodo
                {
                    Periodo = periodo,
                    CarpetaPeriodo = CarpetaPeriodo(codigoPeriodo),
                    RutaPlantilla = PlantillaVigente(),
                    RutaSalida = salida
                });
                resultado = proceso.Resultado;
                leafs = proceso.Leafs;
            }
            catch (CalculoInvalidoException ex) when (ex.Codigo == CodigoError.Validacion)
            {
                // ÚNICA tolerancia declarada: la divergencia de espejo PRE-EXISTENTE (O25/D34/F34),
                // fuera del alcance del Plan 57. El workbook marcado "sin validar" ya está escrito y es
                // el que lee la corrida real. Cualquier otro código (Plantilla, Fuente, ...) NO se
                // tolera: propagaría como rojo de este test.
            }

            var ruta = File.Exists(salida) ? salida : RutaSinValidar(salida);
            Assert.True(File.Exists(ruta), $"La corrida de {codigoPeriodo} no dejó workbook en '{ruta}'.");

            var corrida = new Corrida(
                ruta,
                new ValidacionOracleReader().LeerSnapshots(ruta, periodo),
                resultado,
                leafs);

            Corridas[codigoPeriodo] = corrida;
            return corrida;
        }
    }

    private static ProcesadorPeriodo CrearProcesador() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader(),
            new ValidacionOracleReader());

    /// <summary>Hoja!celda que declara la publicación del EXTEMP del bloque (D{46+aseId}).</summary>
    private static (string Hoja, string Celda) ReferenciaPublicada(string ruta, int aseId)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var formula = Formula(parte, HojaConsolidado, $"D{46 + aseId}");

        var referencia = Regex.Match(formula, @"^'?(?<hoja>[^'!]+)'?!(?<celda>\$?[A-Z]{1,3}\$?\d+)$");
        Assert.True(referencia.Success, $"La publicación del bloque {aseId} no es una referencia única: '{formula}'.");

        return (referencia.Groups["hoja"].Value, referencia.Groups["celda"].Value.Replace("$", string.Empty, StringComparison.Ordinal));
    }

    private static string Formula(WorkbookPart parte, string hoja, string celda)
    {
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheet = ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;
        var cell = worksheet.Descendants<Cell>()
            .First(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));

        return cell.CellFormula?.Text ?? string.Empty;
    }

    /// <summary>Convierte el ancla en un literal (vía literal del oráculo): quita <c>&lt;f&gt;</c>.</summary>
    private static void ConvertirEnLiteral(string ruta, string hoja, string celda, decimal valor)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var cell = BuscarCelda(parte, hoja, celda);
        cell.CellFormula = null;
        cell.CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture));
        cell.DataType = CellValues.Number;
        Guardar(parte, hoja);
    }

    private static void EscribirFormula(string ruta, string hoja, string celda, string formula)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var cell = BuscarCelda(parte, hoja, celda);
        cell.CellFormula = new CellFormula(formula);
        cell.CellValue = new CellValue("0");
        Guardar(parte, hoja);
    }

    /// <summary>Injerto conductual: añade un término falso a la fórmula leída del ancla.</summary>
    private static void AnexarTermino(string ruta, string hoja, string celda, decimal termino)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var cell = BuscarCelda(parte, hoja, celda);
        var formula = cell.CellFormula ?? throw new InvalidOperationException($"El ancla '{celda}' no trae <f>.");
        Assert.False(string.IsNullOrWhiteSpace(formula.Text), $"El ancla '{celda}' es una fórmula compartida sin texto.");

        formula.Text = formula.Text + "+" + termino.ToString(CultureInfo.InvariantCulture);
        formula.FormulaType = null;
        Guardar(parte, hoja);
    }

    private static Cell BuscarCelda(WorkbookPart parte, string hoja, string celda)
    {
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheet = ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;

        return worksheet.Descendants<Cell>()
            .First(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
    }

    private static void Guardar(WorkbookPart parte, string hoja)
    {
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!.Save();
    }

    private static SortedDictionary<string, string> FotoFormulas(string ruta, string hoja)
    {
        var foto = new SortedDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheet = ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;
        foreach (var cell in worksheet.Descendants<Cell>())
        {
            if (cell.CellFormula is not { } formula)
            {
                continue;
            }

            foto[cell.CellReference?.Value ?? "?"] =
                $"{formula.Text}|t={formula.FormulaType?.Value.ToString() ?? string.Empty}|ref={formula.Reference?.Value ?? string.Empty}|si={formula.SharedIndex?.Value.ToString(CultureInfo.InvariantCulture) ?? string.Empty}";
        }

        return foto;
    }

    private sealed class CopiaTemporal : IDisposable
    {
        public CopiaTemporal(string origen)
        {
            var directorio = Path.Combine(Path.GetTempPath(), $"p57-copia-{Environment.ProcessId}-{Guid.NewGuid():N}");
            Directory.CreateDirectory(directorio);
            Ruta = Path.Combine(directorio, Path.GetFileName(origen));
            File.Copy(origen, Ruta);
        }

        public string Ruta { get; }

        public void Dispose()
        {
            try
            {
                var directorio = Path.GetDirectoryName(Ruta);
                if (directorio is not null && Directory.Exists(directorio))
                {
                    Directory.Delete(directorio, recursive: true);
                }
            }
            catch (IOException)
            {
                // best-effort
            }
        }
    }

    private static string Sha256(string ruta)
    {
        using var stream = File.OpenRead(ruta);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    // ── rutas explícitas ──────────────────────────────────────────────────────────────────────

    private static string RaizRepo()
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

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    private static string Ruta(string relativo) => Path.Combine(RaizRepo(), relativo);

    /// <summary>
    /// Plantilla de origen de la corrida publicada: la VIGENTE (v4 desde 2026-10-27). La v2 queda
    /// histórica porque no declara el grupo <c>canal A / empresa EAB</c> del segundo cuadro de R2 que la
    /// fuente de octubre-1 trae en ASE2 (<c>ERR-SIN-SLOT</c> al escribir el detalle).
    /// </summary>
    private static string PlantillaVigente() => Insumos.PlantillaOficialV4();

    private static string CarpetaPeriodo(string codigoPeriodo) => Path.Combine(
        RaizRepo(),
        codigoPeriodo switch
        {
            "2026071" => Path.Combine("Docs", "Prueba Julio-1", "Insumos"),
            "2026072" => Path.Combine("Docs", "Prueba Julio-2", "Insumos"),
            "2026101" => Path.Combine("Docs", "Prueba Octubre-1", "Insumos"),
            _ => throw new ArgumentOutOfRangeException(nameof(codigoPeriodo), codigoPeriodo, "Período sin insumos declarados.")
        });

    private static string RutaSinValidar(string ruta) =>
        Remuneracion.Core.Constants.NombresSalida.RutaSinValidar(ruta);
}
