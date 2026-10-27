using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (§11-bis): el DESGLOSE-DETALLE por componente de <c>Rem. Anticipos R2</c> y
/// <c>Reversion Pagos R4</c> se escribe en AMBAS quincenas, con sus tres niveles —detalle por
/// canal/concepto, totales por canal/empresa y cierre de sección—, y lo escrito se contrasta contra
/// lo que el DOMINIO declara y contra el oráculo administrativo.
///
/// Defecto que cubre (medido en el entregable real de 2026101 — quincena 1): el reader construía el
/// detalle sólo bajo <c>if (esQuincena2)</c>. En la primera quincena el bloque quedaba entero en los
/// literales <c>0</c> de la plantilla, su cierre de cuadro (<c>=E25+E42-K25</c>) daba 0 y la línea de
/// anticipos de <c>CONSOLIDADO!D28:D32</c> se caía: el archivo se contradecía a sí mismo (el
/// <c>DetRetri</c> se escribe desde el dominio y el bloque del que depende estaba vacío).
///
/// Lo que se prueba, sin escribir ni una coordenada (todo sale de la gramática de cierres de la hoja
/// y de las etiquetas que la propia hoja declara):
/// <list type="number">
/// <item>Q1 de 2026101: el gran total de cada bloque coincide, componente por componente, con el que
/// declara el oráculo administrativo del mismo período.</item>
/// <item>Q1 de 2026071: el bloque queda escrito y cuadra contra la fuente en los 5 ASE.</item>
/// <item>Inyección: vaciar el bloque (sólo el detalle, y después el bloque entero) ENROJECE con
/// <c>ERR-ESCRITURA</c> y el workbook no se entrega.</item>
/// </list>
/// </summary>
public sealed class DetalleR2R4EscritoPlan59Tests
{
    private const decimal Tol = 0.5m;

    private static readonly DelimitadorBloquesHoja.HojaDelimitable HojaR2 = DelimitadorBloquesHoja.R2;
    private static readonly DelimitadorBloquesHoja.HojaDelimitable HojaR4 = DelimitadorBloquesHoja.R4;

    /// <summary>Plantilla v4 VIGENTE (declara el grupo <c>A/EAB</c> del 2.º cuadro de R2 en los 5 ASE).</summary>
    private static string PlantillaVigente => Insumos.PlantillaOficialV4();

    /// <summary>Oráculo administrativo de octubre-1: trae los tres niveles del bloque llenos en quincena 1.</summary>
    private static string Oraculo2026101 => Ruta(
        "Docs", "Prueba Octubre-1", "Resultados", "Resultado administrativo", "Remuneracion 202610-1 Total.xlsx");

    [Fact]
    public void Q1_2026101_DetalleR2R4_COINCIDE_CON_EL_ORACULO_ADMINISTRATIVO()
    {
        var periodo = Periodo.Parse("2026101");
        var (leafs, resultado) = ConstruirPeriodo(periodo, Ruta("Docs", "Prueba Octubre-1", "Insumos"));
        var directorio = NuevoDirectorio();
        try
        {
            var salida = Path.Combine(directorio, "salida.xlsx");
            new OpenXmlPlantillaWriter().GenerarWorkbook(PlantillaVigente, salida, resultado, leafs);

            var divergencias = new List<string>();
            foreach (var hoja in new[] { HojaR2, HojaR4 })
            {
                foreach (var leaf in leafs)
                {
                    var aseId = leaf.Ase.Id;
                    var escrito = GranTotalDelBloque(salida, hoja, aseId);
                    var oraculo = GranTotalDelBloque(Oraculo2026101, hoja, aseId);

                    var comparados = 0;
                    foreach (var (etiqueta, valor) in escrito)
                    {
                        if (EsParDebCred(etiqueta) || !oraculo.TryGetValue(etiqueta, out var esperado))
                        {
                            continue;
                        }

                        comparados++;
                        if (Math.Abs(valor - esperado) > Tol)
                        {
                            divergencias.Add($"{hoja.Nombre} ASE {aseId} · '{etiqueta}': escrito={valor} oráculo={esperado} (Δ={valor - esperado}).");
                        }
                    }

                    // Las DOS columnas de Deb/Cred se comparan como PAR: el administrativo clasifica en
                    // 'Deb/Cred-otros' un subconjunto que la fuente trae dentro de 'Deb/Cred' (ASE2 · el
                    // 'Vlr Intereses' de OCCIDENTE: 6.759.294,81 movido de Q a R). Esa clasificación no
                    // está en la fuente y es la MISMA familia del pendiente Plan 58-P2 (columna de
                    // disposición final), declarado NO tocado por el plan: se verifica el TOTAL, que sí es
                    // derivable, y no se inventa la partición.
                    comparados++;
                    var escritoPar = ParDebCred(escrito);
                    var oraculoPar = ParDebCred(oraculo);
                    if (Math.Abs(escritoPar - oraculoPar) > Tol)
                    {
                        divergencias.Add($"{hoja.Nombre} ASE {aseId} · 'Deb/Cred + Deb/Cred-otros': escrito={escritoPar} oráculo={oraculoPar} (Δ={escritoPar - oraculoPar}).");
                    }

                    Assert.True(comparados > 0, $"{hoja.Nombre} ASE {aseId}: no se comparó ningún componente contra el oráculo.");
                }
            }

            Assert.True(divergencias.Count == 0, "El desglose-detalle de Q1 no coincide con el oráculo administrativo:\n" + string.Join("\n", divergencias));

            // El bloque YA NO está vacío: el detalle y los cierres del ASE1 traen monto. Son los tres
            // niveles que el defecto dejaba en el 0 de la plantilla.
            foreach (var hoja in new[] { HojaR2, HojaR4 })
            {
                var niveles = NivelesDelBloque(salida, hoja, leafs[0].Ase.Id, EncabezadoTotal);
                Assert.NotEqual(0m, niveles.Detalle);
                Assert.NotEqual(0m, niveles.Cierres);
                Assert.NotEqual(0m, niveles.GranTotal);
            }
        }
        finally
        {
            BorrarDirectorio(directorio);
        }
    }

    private const string ColumnaDebCred = "Deb/Cred";
    private const string ColumnaDebCredOtros = "Deb/Cred-otros";

    private static bool EsParDebCred(string etiqueta) =>
        string.Equals(etiqueta, ColumnaDebCred, StringComparison.OrdinalIgnoreCase)
        || string.Equals(etiqueta, ColumnaDebCredOtros, StringComparison.OrdinalIgnoreCase);

    private static decimal ParDebCred(IReadOnlyDictionary<string, decimal> valores) =>
        valores.Where(v => EsParDebCred(v.Key)).Sum(v => v.Value);

    [Fact]
    public void Q1_2026071_DetalleR2R4_QUEDA_ESCRITO_Y_CUADRA_CON_LA_FUENTE()
    {
        var periodo = Periodo.Parse("2026071");
        var (leafs, resultado) = ConstruirPeriodo(periodo, Ruta("Docs", "Prueba Julio-1", "Insumos"));
        var directorio = NuevoDirectorio();
        try
        {
            var salida = Path.Combine(directorio, "salida.xlsx");

            // El gate del writer corre DENTRO de GenerarWorkbook: si el bloque quedara vacío o
            // desalineado, esta llamada lanza ERR-ESCRITURA y borra el parcial.
            new OpenXmlPlantillaWriter().GenerarWorkbook(PlantillaVigente, salida, resultado, leafs);

            foreach (var hoja in new[] { HojaR2, HojaR4 })
            {
                foreach (var leaf in leafs)
                {
                    var aseId = leaf.Ase.Id;
                    var escrito = GranTotalDelBloque(salida, hoja, aseId);
                    var esperado = SumaDelDominio(hoja, leaf);

                    var comparados = 0;
                    foreach (var (etiqueta, valor) in esperado)
                    {
                        if (Math.Abs(valor) <= Tol || !escrito.TryGetValue(etiqueta, out var granTotal))
                        {
                            continue;
                        }

                        Assert.True(
                            Math.Abs(granTotal - valor) <= Tol,
                            $"{hoja.Nombre} ASE {aseId} · componente '{etiqueta}': gran total escrito={granTotal} fuente={valor} (Δ={granTotal - valor}).");
                        comparados++;
                    }

                    Assert.True(comparados > 0, $"{hoja.Nombre} ASE {aseId}: el dominio declara montos y no se comparó ningún componente (bloque vacío).");
                }
            }
        }
        finally
        {
            BorrarDirectorio(directorio);
        }
    }

    [Fact]
    public void INYECCION_DetalleForzadoACero_ENROJECE_CON_ERR_ESCRITURA()
    {
        var periodo = Periodo.Parse("2026101");
        var (leafs, resultado) = ConstruirPeriodo(periodo, Ruta("Docs", "Prueba Octubre-1", "Insumos"));
        var directorio = NuevoDirectorio();
        try
        {
            var salida = Path.Combine(directorio, "salida.xlsx");
            new OpenXmlPlantillaWriter().GenerarWorkbook(PlantillaVigente, salida, resultado, leafs);

            // Control: tal como lo dejó el writer, el gate pasa en las dos hojas.
            EjecutarGate(salida, leafs);

            // (a) Sólo el DETALLE en cero: la cadena de agregación del bloque deja de cerrar
            // (detalle 0 contra totales ya escritos) ⇒ ERR-ESCRITURA.
            ForzarACeroNivel(salida, HojaR2, aseId: 2, DelimitadorBloquesHoja.TipoFilaBloqueHoja.Dato);
            var sinDetalle = Assert.Throws<CalculoInvalidoException>(() => EjecutarGate(salida, leafs));
            Assert.Equal(CodigoError.Escritura, sinDetalle.Codigo);
            Assert.Contains("ERR-ESCRITURA", sinDetalle.Message, StringComparison.Ordinal);
            Assert.Contains(HojaR2.Nombre, sinDetalle.Message, StringComparison.Ordinal);
            Assert.Contains("ASE 2", sinDetalle.Message, StringComparison.Ordinal);

            // Se rehace el entregable limpio para la segunda inyección (el gate abortó sobre el archivo
            // ya manipulado y no hay parcial que reutilizar).
            new OpenXmlPlantillaWriter().GenerarWorkbook(PlantillaVigente, salida, resultado, leafs);

            // (b) El BLOQUE ENTERO en cero en R4 (detalle + cierres + gran total): la cadena cierra
            // (0 = 0 = 0) pero la fuente declara monto ⇒ ERR-ESCRITURA por contraste con el dominio.
            ForzarACeroBloqueEntero(salida, HojaR4, aseId: 2);
            var bloqueVacio = Assert.Throws<CalculoInvalidoException>(() => EjecutarGate(salida, leafs));
            Assert.Equal(CodigoError.Escritura, bloqueVacio.Codigo);
            Assert.Contains("ERR-ESCRITURA", bloqueVacio.Message, StringComparison.Ordinal);
            Assert.Contains(HojaR4.Nombre, bloqueVacio.Message, StringComparison.Ordinal);
            Assert.Contains("ASE 2", bloqueVacio.Message, StringComparison.Ordinal);
        }
        finally
        {
            BorrarDirectorio(directorio);
        }
    }

    // ── Lectura por estructura (gramática de cierres + etiquetas de la hoja) ─────────────────────

    private const string EncabezadoTotal = "Total";

    /// <summary>Gran total del bloque del ASE, por etiqueta de componente (leída de la hoja).</summary>
    private static Dictionary<string, decimal> GranTotalDelBloque(string ruta, DelimitadorBloquesHoja.HojaDelimitable hoja, int aseId)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hoja, aseId);
        var filas = FilasDe(workbookPart, hoja);
        var columnas = ColumnasDeValor(workbookPart, filas, bloque);

        var granTotal = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        foreach (var (etiqueta, columna) in columnas)
        {
            granTotal[etiqueta] = ValorDe(workbookPart, filas, bloque.GranTotal, columna);
        }

        return granTotal;
    }

    /// <summary>Σ de las filas de DETALLE que el dominio declara, por etiqueta de componente.</summary>
    private static Dictionary<string, decimal> SumaDelDominio(DelimitadorBloquesHoja.HojaDelimitable hoja, WorkbookLeafInputs leaf)
    {
        var suma = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        void Acumular(string concepto, IReadOnlyDictionary<string, decimal?> valores)
        {
            if (string.IsNullOrWhiteSpace(concepto)
                || string.Equals(concepto.Trim(), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase))
            {
                return; // cierre de la fuente, no detalle
            }

            foreach (var (etiqueta, valor) in valores)
            {
                suma[etiqueta] = suma.GetValueOrDefault(etiqueta) + (valor ?? 0m);
            }
        }

        if (hoja.Nombre == HojaR2.Nombre)
        {
            foreach (var fila in leaf.DetalleR2!.Filas)
            {
                Acumular(fila.D, fila.ValoresPorComponente);
            }
        }
        else
        {
            foreach (var fila in leaf.DetalleR4!.Filas)
            {
                Acumular(fila.C, fila.ValoresPorComponente);
            }
        }

        return suma;
    }

    /// <summary>Σ de los tres niveles del bloque para un componente dado.</summary>
    private static (decimal Detalle, decimal Cierres, decimal GranTotal) NivelesDelBloque(
        string ruta,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        int aseId,
        string etiqueta)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hoja, aseId);
        var filas = FilasDe(workbookPart, hoja);
        var columna = ColumnasDeValor(workbookPart, filas, bloque)
            .First(c => string.Equals(c.Etiqueta, etiqueta, StringComparison.OrdinalIgnoreCase))
            .Columna;

        decimal Sumar(DelimitadorBloquesHoja.TipoFilaBloqueHoja nivel) =>
            bloque.Filas.Where(f => f.Tipo == nivel)
                .Sum(f => ValorDe(workbookPart, filas, f.Fila, columna));

        return (
            Sumar(DelimitadorBloquesHoja.TipoFilaBloqueHoja.Dato),
            Sumar(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreCanal)
                + Sumar(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreEmpresa)
                + Sumar(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion),
            ValorDe(workbookPart, filas, bloque.GranTotal, columna));
    }

    private static Dictionary<int, Row> FilasDe(WorkbookPart workbookPart, DelimitadorBloquesHoja.HojaDelimitable hoja)
    {
        var worksheet = Hoja(workbookPart, hoja.Nombre);
        var tabla = worksheet.Elements<SheetData>().FirstOrDefault()
            ?? throw new InvalidOperationException($"La hoja '{hoja.Nombre}' no tiene SheetData.");
        return tabla.Elements<Row>()
            .Where(f => f.RowIndex is not null)
            .GroupBy(f => (int)f.RowIndex!.Value)
            .ToDictionary(g => g.Key, g => g.First());
    }

    /// <summary>Etiquetas de la fila de encabezados del bloque con la columna que cada una declara.</summary>
    private static IReadOnlyList<(string Etiqueta, string Columna)> ColumnasDeValor(
        WorkbookPart workbookPart,
        IReadOnlyDictionary<int, Row> filas,
        DelimitadorBloquesHoja.BloqueHoja bloque)
    {
        if (!filas.TryGetValue(bloque.EncabezadoColumnas, out var encabezado))
        {
            throw new InvalidOperationException($"No existe la fila de encabezados {bloque.EncabezadoColumnas} de '{bloque.Hoja}'.");
        }

        var columnas = new List<(string, string)>();
        foreach (var celda in encabezado.Elements<Cell>())
        {
            var etiqueta = TextoDe(workbookPart, celda).Trim();
            if (etiqueta.Length == 0)
            {
                continue;
            }

            columnas.Add((etiqueta, ColumnaDe(celda.CellReference?.Value)));
        }

        return columnas;
    }

    private static decimal ValorDe(
        WorkbookPart workbookPart,
        IReadOnlyDictionary<int, Row> filas,
        int numeroFila,
        string columna)
    {
        if (!filas.TryGetValue(numeroFila, out var fila))
        {
            return 0m;
        }

        var referencia = columna + numeroFila.ToString(CultureInfo.InvariantCulture);
        var celda = fila.Elements<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        if (celda is null)
        {
            return 0m;
        }

        var texto = TextoDe(workbookPart, celda);
        return decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor) ? valor : 0m;
    }

    private static string TextoDe(WorkbookPart workbookPart, Cell celda)
    {
        if (celda.DataType?.Value == CellValues.SharedString
            && int.TryParse(celda.CellValue?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice))
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable?
                .Elements<SharedStringItem>().ElementAtOrDefault(indice)?.InnerText ?? string.Empty;
        }

        return celda.CellValue?.InnerText ?? string.Empty;
    }

    private static string ColumnaDe(string? referencia)
    {
        if (string.IsNullOrEmpty(referencia))
        {
            return string.Empty;
        }

        var fin = 0;
        while (fin < referencia.Length && (char.IsLetter(referencia[fin]) || referencia[fin] == '$'))
        {
            fin++;
        }

        return referencia[..fin].Replace("$", string.Empty, StringComparison.Ordinal);
    }

    // ── Inyección (nunca sobre el repo: siempre sobre la salida temporal) ────────────────────────

    private static void ForzarACeroNivel(
        string ruta,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        int aseId,
        DelimitadorBloquesHoja.TipoFilaBloqueHoja nivel)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hoja, aseId);
        var worksheet = Hoja(workbookPart, hoja.Nombre);
        var filas = FilasDe(workbookPart, hoja);
        var columnas = ColumnasDeValor(workbookPart, filas, bloque).Select(c => c.Columna).ToList();

        foreach (var fila in bloque.Filas.Where(f => f.Tipo == nivel))
        {
            foreach (var columna in columnas)
            {
                PonerEnCero(filas, fila.Fila, columna);
            }
        }

        worksheet.Save();
    }

    private static void ForzarACeroBloqueEntero(string ruta, DelimitadorBloquesHoja.HojaDelimitable hoja, int aseId)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hoja, aseId);
        var worksheet = Hoja(workbookPart, hoja.Nombre);
        var filas = FilasDe(workbookPart, hoja);
        var columnas = ColumnasDeValor(workbookPart, filas, bloque).Select(c => c.Columna).ToList();

        foreach (var fila in bloque.Filas.Where(f => f.EsSlot))
        {
            foreach (var columna in columnas)
            {
                PonerEnCero(filas, fila.Fila, columna);
            }
        }

        worksheet.Save();
    }

    private static void PonerEnCero(IReadOnlyDictionary<int, Row> filas, int numeroFila, string columna)
    {
        if (!filas.TryGetValue(numeroFila, out var fila))
        {
            return;
        }

        var referencia = columna + numeroFila.ToString(CultureInfo.InvariantCulture);
        var celda = fila.Elements<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        if (celda is null)
        {
            return;
        }

        celda.CellFormula?.Remove();
        celda.DataType = CellValues.Number;
        celda.CellValue = new CellValue("0");
    }

    private static void EjecutarGate(string ruta, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        OpenXmlPlantillaWriter.VerificarDetalleR2R4(workbookPart, leafs);
    }

    // ── Armado del período con los insumos realmente en disco ───────────────────────────────────

    private static (List<WorkbookLeafInputs> Leafs, ResultadoRemuneracion Resultado) ConstruirPeriodo(
        Periodo periodo,
        string carpetaPeriodo)
    {
        var leafReader = new ExcelDataReaderWorkbookLeafInputReader();
        var lector = new ExcelDataReaderRecaudoReader();
        var leafs = new List<WorkbookLeafInputs>();
        var datos = new List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4, decimal ajustesSfT)>();

        for (var i = 1; i <= 5; i++)
        {
            var carpetaAse = Directory.GetDirectories(carpetaPeriodo)
                .First(d => Path.GetFileName(d).StartsWith($"{i}-", StringComparison.Ordinal));
            var r1 = Fuente(carpetaAse, "Recaudoporcomponente", i);
            var r2 = Fuente(carpetaAse, "RerpoteDetalleSaldosaFavor", i);
            var r4 = FuenteR4(carpetaAse, i);
            var ase = Insumos.Ase(i);

            var leaf = leafReader.LeerLeafInputs(ase, periodo, r1, r2, r4);
            leaf.BalanceSc = leafReader.LeerBalanceSc(ase, periodo, FuenteBalance(carpetaAse, i));

            leafs.Add(leaf);
            datos.Add((ase, lector.LeerR1(r1), lector.LeerR2(r2), lector.LeerR4(r4), 0m));
        }

        return (leafs, new CalculoRemuneracion().CalcularConsolidado(periodo, datos));
    }

    private static string Fuente(string carpetaAse, string prefijo, int aseId)
    {
        var archivo = Directory.EnumerateFiles(carpetaAse, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase))
            ?? throw new FileNotFoundException($"Falta '{prefijo}' del ASE {aseId} en {carpetaAse}.");
        return archivo;
    }

    private static string FuenteR4(string carpetaAse, int aseId) =>
        Fuente(carpetaAse, "Reversi", aseId);

    private static string FuenteBalance(string carpetaAse, int aseId) =>
        Fuente(carpetaAse, "R4-BalanceSubsidioyContribuciones", aseId);

    private static string Ruta(params string[] partes) =>
        Path.Combine([Insumos.Raiz(), .. partes]);

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet.");
    }

    private static string NuevoDirectorio()
    {
        var dir = Path.Combine(Path.GetTempPath(), "p59-detalle-r2r4-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void BorrarDirectorio(string ruta)
    {
        try
        {
            if (Directory.Exists(ruta))
            {
                Directory.Delete(ruta, recursive: true);
            }
        }
        catch
        {
            // best-effort (salida temporal)
        }
    }
}
