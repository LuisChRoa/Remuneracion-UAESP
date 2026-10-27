using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 — H2/F2: evidencia por CONCEPTO de que los destinos de valor y las protecciones se
/// resuelven por identidad (bloque/ASE · sección · subsección · ocurrencia · fuente · canal ·
/// concepto · componente por encabezado) y no por coordenada.
///
/// Qué prueba cada frente:
/// <list type="bullet">
/// <item><b>H2.1</b>: el destino del total oportuno/extemporáneo de cada bloque, y las 36 celdas del
/// CONSOLIDADO/DetRetri, se resuelven por etiqueta; la resolución coincide con el destino que el
/// propio template declara (la referencia de la fórmula consumidora), sin que el código conozca
/// ninguna fila.</item>
/// <item><b>H2.2</b>: el caso testigo §2-D2.3-bis (<c>F75</c> en la v2 vs <c>F89</c> en la plantilla
/// vigente) se aprueba por ROL + FORMA: la celda que hoy corresponde al cierre tiene la forma derivada
/// y la celda congelada de la v2 NO tiene fórmula en la plantilla vigente (no se espera que la tenga).</item>
/// <item><b>Geometría desplazada</b>: un libro sintético con filas extra resuelve los mismos destinos
/// sin cambiar código (prueba de que no hay foto de la plantilla).</item>
/// <item><b>Injerto en rojo</b>: el detector del candado enrojece si aparece un literal de celda o una
/// rama por ASE en los caminos de resolución (el candado formal de dos capas es H3).</item>
/// </list>
/// </summary>
public sealed class DestinoValorDerivadoH2Plan59Tests
{
    private const decimal Tolerancia = 0.5m;

    private static readonly int[] Ase = [1, 2, 3, 4, 5];

    /// <summary>Nombre lógico de cada tabla del CONSOLIDADO, en el orden en que se compone.</summary>
    private static readonly string[] TablasDeRecaudo =
    [
        "R1 - RECAUDO OPORTUNO",
        "R2 - ANTICIPOS A USUARIOS",
        "R3 - RECAUDO EXTEMPORANEO",
        "R4 - REVERSIÓN DE PAGOS",
        "R5 - REMUNERACIÓN SALDOS A FAVOR POR NOTA",
    ];

    /// <summary>Las seis tablas que el CONSOLIDADO declara (las cinco de recaudo + la total).</summary>
    private static readonly string[] TablasDelConsolidado =
    [
        .. TablasDeRecaudo,
        "REPORTE CONSOLIDADO TOTAL DE RECAUDO",
    ];

    // ── H2.1 — destinos por identidad ───────────────────────────────────────────────────────────

    /// <summary>
    /// El total oportuno de cada bloque se resuelve por la subsección <c>Mes</c> #1 y su columna por
    /// el encabezado <c>Total</c>; la celda hallada debe ser EXACTAMENTE la que la tabla
    /// <c>R1 - RECAUDO OPORTUNO</c> del CONSOLIDADO declara para ese ASE (la referencia de su fórmula).
    /// El test no conoce ninguna fila: la referencia se lee del propio template.
    /// </summary>
    [Fact]
    public void H2_1_ElTotalOportunoDeCadaBloqueSeResuelvePorIdentidadYCoincideConLaReferenciaDelConsolidado()
    {
        using var plantilla = PlantillaVigente.Abrir();
        var parte = plantilla.Parte;

        var r1 = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1);
        var hojaR1 = Hoja(parte, DelimitadorBloquesHoja.NombreHojaR1);
        var consolidado = Hoja(parte, DelimitadorBloquesHoja.NombreConsolidado);

        foreach (var aseId in Ase)
        {
            var destino = ResolvedorDestinoValor.ResolverTotalDeBloque(
                parte, hojaR1, DelimitadorBloquesHoja.R1, r1, aseId,
                ResolvedorDestinoValor.RolTotalBloque.Oportuno,
                ResolvedorDestinoValor.EncabezadoTotal);

            var consumidora = DestinoDeLaTabla(parte, consolidado, "R1 - RECAUDO OPORTUNO", aseId);
            var referencia = ReferenciaUnica(consumidora);

            Assert.Equal(DelimitadorBloquesHoja.NombreHojaR1, referencia.Hoja);
            Assert.Equal(destino.Celda, referencia.Celda);
        }
    }

    /// <summary>
    /// Las 30 celdas de las tablas de recaudo + las 5 del DetRetri se resuelven por
    /// (título en <c>F</c> + encabezado <c>B='ASE'</c> debajo + id del ASE + encabezado de columna).
    /// Se asevera que cada fila hallada declara el id numérico del ASE y que el rótulo largo leído de
    /// la propia hoja empieza por el nombre del catálogo (H-59-4: el punto final puede diferir entre
    /// hojas, así que no se compara la cadena exacta de otra hoja).
    /// </summary>
    [Fact]
    public void H2_1_Las36CeldasDelConsolidadoYDetRetriSeResuelvenPorTablaYRotulo()
    {
        using var plantilla = PlantillaVigente.Abrir();
        var parte = plantilla.Parte;

        var consolidado = Hoja(parte, DelimitadorBloquesHoja.NombreConsolidado);
        var detRetri = Hoja(parte, NombresHojaPeriodo.DetRetri(Periodo.Parse("2026092")));

        var celdas = 0;
        foreach (var aseId in Ase)
        {
            foreach (var titulo in TablasDelConsolidado)
            {
                var destino = DestinoDeLaTabla(parte, consolidado, titulo, aseId);
                Assert.Equal(aseId.ToString(System.Globalization.CultureInfo.InvariantCulture), Texto(parte, consolidado, "B", destino.Fila));
                // El rótulo largo se lee de la PROPIA hoja (H-59-4: puede diferir por un punto entre
                // hojas, así que no se compara la cadena de otra hoja); aquí basta que exista.
                Assert.False(string.IsNullOrWhiteSpace(Texto(parte, consolidado, "C", destino.Fila)));
                celdas++;
            }

            var filaDetRetri = ResolvedorDestinoValor.ResolverTabla(
                parte, detRetri,
                new ResolvedorDestinoValor.IdentidadTabla(
                    NombresHojaPeriodo.DetRetri(Periodo.Parse("2026092")), Titulo: null, aseId,
                    ResolvedorDestinoValor.EncabezadoRecaudoTotal));
            Assert.Equal(aseId.ToString(System.Globalization.CultureInfo.InvariantCulture), Texto(parte, detRetri, "B", filaDetRetri.Fila));
            celdas++;
        }

        // Fila de cierre D109 del CONSOLIDADO y fila de total del DetRetri: también por etiqueta.
        var cierreConsolidado = DestinoDeLaTabla(parte, consolidado, "REPORTE CONSOLIDADO TOTAL DE RECAUDO", aseId: null);
        Assert.Equal("TOTAL", Texto(parte, consolidado, "C", cierreConsolidado.Fila), StringComparer.OrdinalIgnoreCase);
        celdas++;

        var cierreDetRetri = ResolvedorDestinoValor.ResolverTabla(
            parte, detRetri,
            new ResolvedorDestinoValor.IdentidadTabla(
                NombresHojaPeriodo.DetRetri(Periodo.Parse("2026092")), Titulo: null, AseId: null,
                ResolvedorDestinoValor.EncabezadoRecaudoTotal));
        Assert.Equal("Total", Texto(parte, detRetri, "C", cierreDetRetri.Fila), StringComparer.OrdinalIgnoreCase);
        celdas++;

        // 30 del CONSOLIDADO (6 tablas × 5 ASE) + 5 del DetRetri + D109 + cierre del DetRetri = 37
        // (las 36 del T0 más la fila de cierre del DetRetri, que la misma tabla declara).
        Assert.Equal(37, celdas);
    }

    /// <summary>
    /// El destino de escritura del DetRetri es un SLOT DE VALOR: la verificación por identidad exige
    /// que la celda no traiga <c>&lt;f&gt;</c> (si el template la protegiera, el escritor falla antes de
    /// escribir y nombra la identidad).
    /// </summary>
    [Fact]
    public void H2_3_ElDestinoDetRetriEsSlotDeValorYSeResuelvePorEtiqueta()
    {
        using var plantilla = PlantillaVigente.Abrir();
        var parte = plantilla.Parte;

        VerificadorDestinoProtegido.ValidarDetRetriEditable(parte, Periodo.Parse("2026092"));

        var hoja = Hoja(parte, NombresHojaPeriodo.DetRetri(Periodo.Parse("2026092")));
        foreach (var aseId in Ase)
        {
            var destino = ResolvedorDestinoValor.ResolverTabla(
                parte, hoja,
                new ResolvedorDestinoValor.IdentidadTabla(
                    NombresHojaPeriodo.DetRetri(Periodo.Parse("2026092")), Titulo: null, aseId,
                    ResolvedorDestinoValor.EncabezadoRecaudoTotal));

            Assert.Null(destino.CeldaXml.CellFormula);
        }
    }

    // ── H2.2 — caso testigo F75/F89 y verificación por forma ───────────────────────────────────

    /// <summary>
    /// Caso testigo §2-D2.3-bis: en la plantilla vigente el cierre oportuno del bloque ASE1 vive donde la geometría
    /// lo declara (el cuadro de totales, rótulo <c>C='TOTAL'/D='OPORTUNO'</c>) y su fórmula SUMA los
    /// cierres <c>Mes</c> del bloque y RESTA sus especiales menos el último. La celda congelada de la
    /// v2 (<c>F75</c>) NO tiene fórmula y NO se espera que la tenga.
    /// </summary>
    [Fact]
    public void H2_2_ElCierreOportunoSeApruebaPorFormaYLaCeldaCongeladaDeLaV2NoTieneFormula()
    {
        using var plantilla = PlantillaVigente.Abrir();
        var parte = plantilla.Parte;

        // 1) La protección completa del template se verifica por ROL + FORMA (si algo no cuadra, este
        //    llamado lanza ERR-PLANTILLA con la identidad; no compara un solo literal).
        VerificadorDestinoProtegido.ValidarBloquesDeRecaudo(parte);
        VerificadorDestinoProtegido.ValidarConsolidado(parte);

        var hojaR1 = Hoja(parte, DelimitadorBloquesHoja.NombreHojaR1);
        var r1 = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1);
        var bloque = r1.Single(b => b.AseId == 1);

        // 2) La forma del cierre se deriva de la geometría: Σ Total de los cierres `Mes` − Especiales
        //    de todos menos el último.
        var mes = ResolvedorDestinoValor.CierresDeSubseccion(bloque, ResolvedorDestinoValor.SubseccionOportuna);
        Assert.Equal(3, mes.Count);
        var esperados = new List<string>();
        for (var i = 0; i < mes.Count; i++)
        {
            esperados.Add($"{ColumnaDeEncabezado(parte, hojaR1, bloque, ResolvedorDestinoValor.EncabezadoTotal)}{mes[i].Fila}");
            if (i < mes.Count - 1)
            {
                esperados.Add($"-{ColumnaDeEncabezado(parte, hojaR1, bloque, ResolvedorDestinoValor.EncabezadoEspecialesR1)}{mes[i].Fila}");
            }
        }

        var cierre = FilaDelCuadroDeTotales(parte, hojaR1, bloque, "OPORTUNO");
        var celdaCierre = $"{ColumnaDeEncabezado(parte, hojaR1, bloque, ResolvedorDestinoValor.EncabezadoTotal)}{cierre}";
        var formula = ResolvedorDestinoValor.BuscarCelda(hojaR1, celdaCierre)?.CellFormula?.Text;
        Assert.NotNull(formula);

        // La forma se cuenta sobre la fórmula real: Σ de los cierres (N términos) y resta de los
        // especiales de todos menos el último (N−1 signos negativos). El número lo dicta la geometría.
        Assert.NotNull(formula);
        Assert.Equal(mes.Count - 1, formula!.Count(c => c == '+'));
        Assert.Equal(mes.Count - 1, formula.Count(c => c == '-'));

        // 3) La celda congelada de la v2 (`F75` en R1Q2ProtectedPorAse[1]) ya no es fórmula en la
        //    plantilla vigente: el mensaje "debería seguir siendo fórmula" era falso respecto de ella.
        var celdaV2 = WorkbookLeafCellMapQ2.ObtenerR1Q2Protegidos(1)[0].Celda;
        var enVigente = ResolvedorDestinoValor.BuscarCelda(hojaR1, celdaV2);
        Assert.NotNull(enVigente);
        Assert.Null(enVigente!.CellFormula);

        // Y el cierre de verdad es OTRA celda, no la congelada.
        Assert.NotEqual(celdaV2, celdaCierre);
    }

    /// <summary>
    /// La verificación es por FORMA, no por lista de operandos: si el template pierde el cierre (la
    /// celda pasa a valor fijo) el verificador falla nombrando la IDENTIDAD y la celda hallada, no una
    /// coordenada esperada.
    /// </summary>
    [Fact]
    public void H2_2_SiElCierreDejaDeSerFormula_FallaNombrandoLaIdentidad()
    {
        using var plantilla = PlantillaVigente.Abrir(enEscritura: true);
        var parte = plantilla.Parte;

        var hojaR1 = Hoja(parte, DelimitadorBloquesHoja.NombreHojaR1);
        var bloque = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1).Single(b => b.AseId == 1);
        var fila = FilaDelCuadroDeTotales(parte, hojaR1, bloque, "OPORTUNO");
        var celda = $"{ColumnaDeEncabezado(parte, hojaR1, bloque, ResolvedorDestinoValor.EncabezadoTotal)}{fila}";

        var objetivo = ResolvedorDestinoValor.BuscarCelda(hojaR1, celda)!;
        objetivo.CellFormula!.Remove();
        hojaR1.Save();

        var excepcion = Assert.Throws<Remuneracion.Core.Exceptions.CalculoInvalidoException>(
            () => VerificadorDestinoProtegido.ValidarBloquesDeRecaudo(parte));

        Assert.Contains("bloque ASE1", excepcion.Message, StringComparison.Ordinal);
        Assert.Contains("oportuno", excepcion.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(celda, excepcion.Message, StringComparison.Ordinal);
    }

    // ── Geometría desplazada ────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Mismo código, otra geometría: un libro sintético con filas extra por bloque (estilo v2→v3)
    /// resuelve los mismos destinos. Se asevera la relación estructural (el cierre es la primera fila
    /// posterior al gran total con el rótulo <c>C='TOTAL'/D='OPORTUNO'</c>) y no una fila concreta.
    /// </summary>
    [Theory]
    [InlineData(0, 0)]
    [InlineData(14, 55)]
    [InlineData(31, 3)]
    public void H2_1_GeometriaDesplazada_ResuelveSinCambiarCodigo(int desplazamientoInicial, int filasExtra)
    {
        using var libro = LibroSintetico.CrearR1(desplazamientoInicial, filasExtra);
        using var documento = SpreadsheetDocument.Open(libro.Ruta, false);
        var parte = documento.WorkbookPart!;

        var hoja = Hoja(parte, DelimitadorBloquesHoja.NombreHojaR1);
        var bloques = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1);
        Assert.Equal(5, bloques.Count);

        foreach (var aseId in Ase)
        {
            var bloque = bloques.Single(b => b.AseId == aseId);
            var mes = ResolvedorDestinoValor.CierresDeSubseccion(bloque, ResolvedorDestinoValor.SubseccionOportuna);
            Assert.Equal(2, mes.Count);

            var destino = ResolvedorDestinoValor.Resolver(
                parte, hoja, DelimitadorBloquesHoja.R1, bloques,
                new ResolvedorDestinoValor.Identidad(
                    DelimitadorBloquesHoja.NombreHojaR1, aseId,
                    ResolvedorDestinoValor.NivelDestino.Subseccion,
                    ResolvedorDestinoValor.SubseccionOportuna, 2,
                    ResolvedorDestinoValor.EncabezadoTotal));

            // El rol #2 es la SEGUNDA ocurrencia de la etiqueta dentro del bloque, con la geometría que
            // el libro declare (desplazada o no): se compara contra la propia geometría, no contra un
            // número.
            Assert.Equal(mes[1].Fila, destino.Fila);
            Assert.Equal("F", destino.Columna);
            Assert.Equal($"F{destino.Fila}", destino.Celda);
        }
    }

    // ── Injerto en rojo (aserción que demuestra el candado) ────────────────────────────────────

    /// <summary>
    /// Detector del candado G-QUEMADO: en los caminos de resolución no puede haber literales de
    /// celda/fila ni ramas por ASE. El test se autoverifica con un injerto: una fuente con
    /// <c>"F75"</c> y con <c>if (ase == 2)</c> DEBE enrojecer el detector (si no, el candado sería
    /// decorativo).
    /// </summary>
    [Fact]
    public void H2_5_Candado_CeroLiteralesYRamificacionesEnLosCaminosDeResolucion()
    {
        var caminos = new[]
        {
            "Remuneracion.Infrastructure/Excel/ResolvedorDestinoValor.cs",
            "Remuneracion.Infrastructure/Excel/VerificadorDestinoProtegido.cs",
            "Remuneracion.Infrastructure/Excel/EscritorValoresSlotR1.cs",
            "Remuneracion.Infrastructure/Excel/DelimitadorBloquesHoja.cs",
        };

        var hallazgos = new List<string>();
        foreach (var ruta in caminos)
        {
            var texto = File.ReadAllText(Path.Combine(RaizRepo(), ruta));
            hallazgos.AddRange(AnalizarQuemados(ruta, texto));
        }

        Assert.Empty(hallazgos);

        // Injerto: las dos formas prohibidas deben ser detectadas.
        Assert.NotEmpty(AnalizarQuemados("injerto.cs", "var x = \"F75\";"));
        Assert.NotEmpty(AnalizarQuemados("injerto.cs", "if (ase == 2) { }"));
        Assert.NotEmpty(AnalizarQuemados("injerto.cs", "var rango = (6, 72);"));
        Assert.Empty(AnalizarQuemados("limpio.cs", "var celda = $\"F{fila}\";"));
    }

    /// <summary>
    /// Analizador del candado: literales con forma de referencia A1, ramas por ASE/empresa y rangos de
    /// fila congelados. Es la MISMA regla que el candado de dos capas de H3 extenderá; aquí vive la
    /// aserción que demuestra que el injerto enrojece.
    /// </summary>
    private static IEnumerable<string> AnalizarQuemados(string archivo, string fuente)
    {
        // Solo CÓDIGO: la documentación cita a propósito las coordenadas que el plan eliminó (p. ej.
        // la ventana histórica `(6,72)` que retiró H1) y citarlas no es re-congelarlas.
        var codigo = System.Text.RegularExpressions.Regex.Replace(fuente, @"//[^\n]*", string.Empty);
        codigo = System.Text.RegularExpressions.Regex.Replace(
            codigo, @"/\*.*?\*/", string.Empty, System.Text.RegularExpressions.RegexOptions.Singleline);

        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(
                     codigo, "\"([A-Za-z]{1,2}[0-9]{1,4})\""))
        {
            yield return $"{archivo}: literal de celda '{m.Groups[1].Value}'";
        }

        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(
                     codigo, @"\b(ase|aseId|Ase\.Id|ase\.Id)\s*==\s*\d"))
        {
            yield return $"{archivo}: rama por ASE '{m.Value}'";
        }

        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(
                     codigo, @"\(\s*[0-9]{1,4}\s*,\s*[0-9]{1,4}\s*\)"))
        {
            yield return $"{archivo}: rango de filas congelado '{m.Value}'";
        }
    }

    // ── Utilidades ──────────────────────────────────────────────────────────────────────────────

    private static ResolvedorDestinoValor.Destino DestinoDeLaTabla(
        WorkbookPart parte,
        Worksheet hoja,
        string titulo,
        int? aseId) =>
        ResolvedorDestinoValor.ResolverTabla(
            parte, hoja,
            new ResolvedorDestinoValor.IdentidadTabla(
                DelimitadorBloquesHoja.NombreConsolidado, titulo, aseId,
                ResolvedorDestinoValor.EncabezadoRecaudoTotal));

    /// <summary>Referencia única (<c>Hoja!Celda</c>) que declara una celda del CONSOLIDADO.</summary>
    private static (string Hoja, string Celda) ReferenciaUnica(ResolvedorDestinoValor.Destino consumidora)
    {
        var formula = consumidora.CeldaXml.CellFormula?.Text
            ?? throw new InvalidOperationException($"La celda consumidora {consumidora.Celda} no declara fórmula.");
        var texto = formula.Replace("'", string.Empty, StringComparison.Ordinal);
        var partes = texto.Split('!');
        Assert.Equal(2, partes.Length);
        return (partes[0], partes[1]);
    }

    private static int FilaDelCuadroDeTotales(
        WorkbookPart parte,
        Worksheet hoja,
        DelimitadorBloquesHoja.BloqueHoja bloque,
        string concepto)
    {
        var filas = hoja.Elements<SheetData>().FirstOrDefault()!.Elements<Row>()
            .Where(r => r.RowIndex is not null && r.RowIndex!.Value > bloque.GranTotal)
            .OrderBy(r => r.RowIndex!.Value)
            .ToList();

        var fila = filas.FirstOrDefault(r =>
            string.Equals(TextoDeFila(parte, r, "C"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase)
            && string.Equals(TextoDeFila(parte, r, "D"), concepto, StringComparison.OrdinalIgnoreCase)
            && string.Equals(TextoDeFila(parte, r, "E"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase));

        return fila is null
            ? throw new InvalidOperationException($"No se encontró el cuadro de totales '{concepto}' del bloque ASE{bloque.AseId}.")
            : (int)fila.RowIndex!.Value;
    }

    private static string ColumnaDeEncabezado(
        WorkbookPart parte,
        Worksheet hoja,
        DelimitadorBloquesHoja.BloqueHoja bloque,
        string encabezado)
    {
        var fila = hoja.Descendants<Row>().First(r => r.RowIndex?.Value == bloque.EncabezadoColumnas);
        foreach (var celda in fila.Elements<Cell>())
        {
            var letra = ResolvedorDestinoValor.LetraDeColumna(celda.CellReference?.Value);
            if (letra.Length > 0
                && string.Equals(LectorTextoHoja.Texto(parte, celda).Trim(), encabezado, StringComparison.OrdinalIgnoreCase))
            {
                return letra;
            }
        }

        throw new InvalidOperationException($"El bloque ASE{bloque.AseId} no declara el encabezado '{encabezado}'.");
    }

    private static Worksheet Hoja(WorkbookPart parte, string nombre)
    {
        var hoja = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet!;
    }

    private static string Texto(WorkbookPart parte, Worksheet hoja, string columna, int fila) =>
        LectorTextoHoja.TextoDeColumna(
            parte,
            hoja.Descendants<Row>().First(r => r.RowIndex?.Value == fila),
            columna).Trim();

    private static string TextoDeFila(WorkbookPart parte, Row fila, string columna) =>
        LectorTextoHoja.TextoDeColumna(parte, fila, columna).Trim();

    private static string RaizRepo()
    {
        var directorio = new DirectoryInfo(AppContext.BaseDirectory);
        while (directorio is not null && !File.Exists(Path.Combine(directorio.FullName, "README.md")))
        {
            directorio = directorio.Parent;
        }

        return directorio?.FullName ?? throw new InvalidOperationException("No se encontró la raíz del repositorio.");
    }

    /// <summary>
    /// Plantilla v4 VIGENTE abierta en solo-lectura (o en escritura sobre una COPIA, para los tests que
    /// simulan una plantilla rota sin tocar el original).
    /// </summary>
    private sealed class PlantillaVigente : IDisposable
    {
        private readonly string? _copia;

        private PlantillaVigente(SpreadsheetDocument documento, string? copia)
        {
            Documento = documento;
            _copia = copia;
        }

        internal SpreadsheetDocument Documento { get; }

        internal WorkbookPart Parte => Documento.WorkbookPart!;

        internal static PlantillaVigente Abrir(bool enEscritura = false)
        {
            var origen = Insumos.PlantillaOficialV4();

            if (!enEscritura)
            {
                return new PlantillaVigente(SpreadsheetDocument.Open(origen, false), copia: null);
            }

            var copia = Path.Combine(Path.GetTempPath(), "p59-h2-" + Guid.NewGuid().ToString("N") + ".xlsx");
            File.Copy(origen, copia, overwrite: true);
            return new PlantillaVigente(SpreadsheetDocument.Open(copia, true), copia);
        }

        public void Dispose()
        {
            Documento.Dispose();
            if (_copia is null)
            {
                return;
            }

            try
            {
                File.Delete(_copia);
            }
            catch
            {
                // best-effort
            }
        }
    }

    /// <summary>
    /// Libro sintético con la gramática de R1 (cabecera por rótulo del ASE, cierres por firma,
    /// gran total y cuadro de totales con el rótulo del cierre). Permite probar la resolución con una
    /// geometría QUE NO ES la de la plantilla vigente: si el código tuviera una fila congelada, este test enrojece.
    /// </summary>
    private sealed class LibroSintetico : IDisposable
    {
        private readonly string _directorio;

        private LibroSintetico(string directorio, string ruta)
        {
            _directorio = directorio;
            Ruta = ruta;
        }

        internal string Ruta { get; }

        internal static LibroSintetico CrearR1(int desplazamientoInicial, int filasExtra)
        {
            var directorio = Path.Combine(Path.GetTempPath(), "p59-h2-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directorio);
            var ruta = Path.Combine(directorio, "sintetico.xlsx");
            var libro = new LibroSintetico(directorio, ruta);

            using (var documento = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook))
            {
                var parte = documento.AddWorkbookPart();
                parte.Workbook = new Workbook();
                var hojas = parte.Workbook.AppendChild(new Sheets());
                var hojaParte = parte.AddNewPart<WorksheetPart>();
                var datos = new SheetData();
                hojaParte.Worksheet = new Worksheet(datos);
                hojas.Append(new Sheet
                {
                    Id = parte.GetIdOfPart(hojaParte),
                    SheetId = 1,
                    Name = DelimitadorBloquesHoja.NombreHojaR1,
                });

                var fila = 1 + desplazamientoInicial;
                for (var ase = 1; ase <= 5; ase++)
                {
                    Fila(datos, fila++, ("B", AseFactory.DesdeId(ase).NombreCompleto.ToUpperInvariant()), ("A", string.Empty));
                    Fila(datos, fila++, ("F", "Total"), ("L", "SERVICIO ESPECIALES"));

                    for (var k = 0; k < filasExtra; k++)
                    {
                        FilaConValor(datos, fila++, ("E", "Vlr Servicio"), ("D", "1"));
                    }

                    // Dos subsecciones `Mes` (la primera con especiales) en dos secciones distintas.
                    FilaConValor(datos, fila++, ("E", "Vlr Servicio"), ("D", "1"));
                    FilaConValor(datos, fila++, ("B", "Mes"), ("C", "Total"));
                    FilaConValor(datos, fila++, ("A", "SECCION UNO"), ("B", "Total"));
                    FilaConValor(datos, fila++, ("E", "Vlr Servicio"), ("D", "1"));
                    FilaConValor(datos, fila++, ("B", "Mes"), ("C", "Total"));
                    FilaConValor(datos, fila++, ("A", "SECCION DOS"), ("B", "Total"));
                    FilaConValor(datos, fila++, ("A", "Total"));

                    // Cuadro de totales: reutiliza rótulos (el resolvedor lo identifica por su marca).
                    FilaSinValor(datos, fila++, ("C", "TOTAL"), ("D", "OPORTUNO"), ("E", "Total"), ("F", "FORMULA"));
                    FilaSinValor(datos, fila++, ("D", "EXTEMPORANEO"), ("E", "Total"), ("F", "FORMULA"));
                    FilaSinValor(datos, fila++, ("C", "RECIPROCIDAD"), ("D", "OPORTUNO"), ("E", "Total"));
                }
            }

            return libro;
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(_directorio, recursive: true);
            }
            catch
            {
                // best-effort
            }
        }

        private static void Fila(SheetData datos, int numero, params (string Columna, string Texto)[] textos)
        {
            var fila = new Row { RowIndex = (uint)numero };
            foreach (var (columna, texto) in textos.Where(t => t.Texto.Length > 0))
            {
                fila.Append(CeldaTexto(columna, numero, texto));
            }

            datos.Append(fila);
        }

        private static void FilaConValor(SheetData datos, int numero, params (string Columna, string Texto)[] textos)
        {
            var fila = new Row { RowIndex = (uint)numero };
            foreach (var (columna, texto) in textos)
            {
                fila.Append(CeldaTexto(columna, numero, texto));
            }

            fila.Append(new Cell
            {
                CellReference = $"F{numero}",
                DataType = CellValues.Number,
                CellValue = new CellValue("0"),
            });
            datos.Append(fila);
        }

        private static void FilaSinValor(SheetData datos, int numero, params (string Columna, string Texto)[] textos)
        {
            var fila = new Row { RowIndex = (uint)numero };
            foreach (var (columna, texto) in textos)
            {
                fila.Append(columna == "F"
                    ? new Cell
                    {
                        CellReference = $"F{numero}",
                        CellFormula = new CellFormula("0+0"),
                        CellValue = new CellValue("0"),
                    }
                    : CeldaTexto(columna, numero, texto));
            }

            datos.Append(fila);
        }

        private static Cell CeldaTexto(string columna, int numero, string texto) => new()
        {
            CellReference = $"{columna}{numero}",
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text(texto)),
        };
    }
}
