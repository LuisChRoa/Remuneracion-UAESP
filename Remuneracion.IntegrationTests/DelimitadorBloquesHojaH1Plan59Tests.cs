using DocumentFormat.OpenXml;
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
/// Plan 59 (H1/F1 — H1.1/H1.3, S2.1/S2.2): evidencia de la <b>geometría DERIVADA</b> de los bloques
/// de ASE sobre la plantilla v3 vigente (<c>SHA256 8199F864…2B1295</c>).
///
/// <para>
/// El fixture prueba tres cosas y ninguna más:
/// </para>
/// <list type="number">
/// <item><b>Resolución correcta</b>: el delimitador devuelve, para las 3 hojas × 5 ASE, exactamente
/// los rangos que el T0 midió por bytes (cabecera de bloque y gran total), y la clasificación por
/// firma de cada fila (sección, subsección, empresa, canal) coincide con el vocabulario de cierres
/// del T0 §2.</item>
/// <item><b>El cuadro resumen queda FUERA</b>: debajo del gran total el cuadro reutiliza firmas de
/// cierre (R2 <c>C='OPORTUNO'/D='Total'</c>; R4 <c>B='OPORTUNO'/C='Total'</c>); ninguna fila de ese
/// cuadro pertenece a un bloque.</item>
/// <item><b>Geometría desplazada</b>: un workbook sintético con filas extra y bloques de distinta
/// longitud resuelve con el MISMO código (cero constantes de fila en producción).</item>
/// </list>
///
/// <para>
/// Los números medidos viven AQUÍ (aserción de test = evidencia, nunca constante de producción) y
/// llevan su derivación: hoja + fila del workbook v3 del disco.
/// </para>
/// </summary>
public sealed class DelimitadorBloquesHojaH1Plan59Tests
{
    // ── Geometría medida en H0 (§0.4-b, §2 del T0) sobre la v3 ────────────────────────────────
    // (Ase, fila de la cabecera de bloque = B con el rótulo del ASE, fila del gran total A='Total').
    private static readonly (int Ase, int Cabecera, int GranTotal)[] BloquesR1 =
        [(1, 4, 86), (2, 122, 279), (3, 356, 424), (4, 517, 674), (5, 725, 807)];

    private static readonly (int Ase, int Cabecera, int GranTotal)[] BloquesR2 =
        [(1, 1, 43), (2, 75, 147), (3, 196, 231), (4, 317, 389), (5, 441, 483)];

    private static readonly (int Ase, int Cabecera, int GranTotal)[] BloquesR4 =
        [(1, 1, 21), (2, 97, 117), (3, 198, 212), (4, 235, 255), (5, 350, 370)];

    /// <summary>
    /// Rango físico medido en H0 por bloque = [cabecera … (cabecera del ASE siguiente − 1)], con el
    /// último bloque cerrando en la última fila del <c>dimension</c>. Se aserta como DATO de la
    /// evidencia: el bloque de DATOS termina antes (en el gran total) porque el hueco contiene el
    /// cuadro resumen (§2.5 del T0).
    /// </summary>
    private static readonly (int Ase, int InicioRegion, int FinRegion)[] RegionesR1 =
        [(1, 4, 121), (2, 122, 355), (3, 356, 516), (4, 517, 724), (5, 725, 843)];

    private static readonly (int Ase, int InicioRegion, int FinRegion)[] RegionesR2 =
        [(1, 1, 74), (2, 75, 195), (3, 196, 316), (4, 317, 440), (5, 441, 527)];

    private static readonly (int Ase, int InicioRegion, int FinRegion)[] RegionesR4 =
        [(1, 1, 96), (2, 97, 197), (3, 198, 234), (4, 235, 349), (5, 350, 394)];

    // ── Cierres declarados por bloque en la v3 (T0 §2.1/§2.2/§2.3) ────────────────────────────
    private static readonly int[] CanalesR1 = [24, 55, 20, 55, 24];
    private static readonly int[] CanalesR2 = [12, 24, 10, 24, 12];
    private static readonly int[] EmpresasR1 = [14, 15, 10, 15, 14];
    private static readonly int[] EmpresasR2 = [8, 8, 6, 8, 8];
    private static readonly int[] EmpresasR4 = [3, 3, 2, 3, 3];
    private static readonly int[] SeccionesR1 = [3, 3, 3, 3, 3];
    private static readonly int[] SubseccionesR1 = [5, 5, 5, 5, 5];
    private static readonly int[] SeccionesR2 = [2, 2, 2, 2, 2];
    private static readonly int[] SeccionesR4 = [6, 6, 4, 6, 6];

    /// <summary>Cierres de SECCIÓN por bloque (fila @ etiqueta), T0 §2.1/§2.2/§2.3.</summary>
    private static readonly (int Fila, string Etiqueta)[][] CierresSeccionR1 =
    [
        [(17, "AFaseo"), (56, "Componente"), (85, "Subs/Cont")],
        [(150, "AFaseo"), (225, "Componente"), (278, "Subs/Cont")],
        [(369, "AFaseo"), (400, "Componente"), (423, "Subs/Cont")],
        [(545, "AFaseo"), (620, "Componente"), (673, "Subs/Cont")],
        [(738, "AFaseo"), (777, "Componente"), (806, "Subs/Cont")],
    ];

    // ═══════════════════════════════════════════════════════════════════════════════════════════
    // H1.1 / H1.3 — las 3 hojas × 5 bloques contra la geometría medida en H0
    // ═══════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H1_1_R1_LosCincoBloquesResuelvenLaGeometriaMedidaEnH0()
    {
        var bloques = Delimitar(DelimitadorBloquesHoja.R1);

        AssertGeometria(bloques, BloquesR1, RegionesR1, DelimitadorBloquesHoja.R1);
        Assert.Equal(SeccionesR1, bloques.Select(b => b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion)).ToArray());
        Assert.Equal(SubseccionesR1, bloques.Select(b => b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSubseccion)).ToArray());
        Assert.Equal(CanalesR1, bloques.Select(b => b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreCanal)).ToArray());
        Assert.Equal(EmpresasR1, bloques.Select(b => b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreEmpresa)).ToArray());
    }

    [Fact]
    public void H1_3_R2_LosCincoBloquesResuelvenLaGeometriaMedidaEnH0_SinSubseccion()
    {
        var bloques = Delimitar(DelimitadorBloquesHoja.R2);

        AssertGeometria(bloques, BloquesR2, RegionesR2, DelimitadorBloquesHoja.R2);
        Assert.Equal(SeccionesR2, bloques.Select(b => b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion)).ToArray());
        Assert.Equal(CanalesR2, bloques.Select(b => b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreCanal)).ToArray());
        Assert.Equal(EmpresasR2, bloques.Select(b => b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreEmpresa)).ToArray());
        Assert.All(bloques, b => Assert.Equal(0, b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSubseccion)));
        Assert.False(DelimitadorBloquesHoja.R2.DeclaraSubseccion);
        Assert.True(DelimitadorBloquesHoja.R2.DeclaraCanal);
    }

    [Fact]
    public void H1_3_R4_LosCincoBloquesResuelvenLaGeometriaMedidaEnH0_SinCanalNiSubseccion()
    {
        var bloques = Delimitar(DelimitadorBloquesHoja.R4);

        AssertGeometria(bloques, BloquesR4, RegionesR4, DelimitadorBloquesHoja.R4);
        Assert.Equal(SeccionesR4, bloques.Select(b => b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion)).ToArray());
        Assert.Equal(EmpresasR4, bloques.Select(b => b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreEmpresa)).ToArray());
        Assert.All(bloques, b =>
        {
            Assert.Equal(0, b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreCanal));
            Assert.Equal(0, b.Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSubseccion));
        });
        Assert.False(DelimitadorBloquesHoja.R4.DeclaraCanal);
        Assert.False(DelimitadorBloquesHoja.R4.DeclaraSubseccion);
    }

    /// <summary>
    /// H1.1 (§2-D1): la identidad del bloque es el RÓTULO del ASE leído de la columna B, resuelto
    /// contra el catálogo de dominio — no un índice ni una fila. Y los cierres de sección de R1
    /// caen exactamente donde el T0 los midió (clasificación por firma, no por posición).
    /// </summary>
    [Fact]
    public void H1_1_LaIdentidadEsElRotuloDelCatalogo_YLosCierresDeSeccionSonLosMedidos()
    {
        var bloques = Delimitar(DelimitadorBloquesHoja.R1);

        for (var i = 0; i < bloques.Count; i++)
        {
            var esperado = AseFactory.DesdeId(bloques[i].AseId).NombreCompleto;
            Assert.True(
                SinonimosEmpresaR1.SonMismaEmpresa(bloques[i].Rotulo, esperado),
                $"El bloque {i + 1} declara el rótulo '{bloques[i].Rotulo}' y no corresponde al ASE {bloques[i].AseId} ('{esperado}').");

            var cierres = bloques[i].Cierres(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion);
            Assert.Equal(CierresSeccionR1[i].Length, cierres.Count);
            for (var k = 0; k < cierres.Count; k++)
            {
                Assert.Equal(CierresSeccionR1[i][k].Fila, cierres[k].Fila);
                Assert.Equal(CierresSeccionR1[i][k].Etiqueta, cierres[k].Etiquetas.A.Trim());
            }
        }
    }

    /// <summary>
    /// H1.1 (trampa medida en H0 §2.5): el cuadro resumen de debajo del gran total reutiliza las
    /// firmas de cierre. El bloque corta en el gran total, así que NINGUNA fila del resumen entra en
    /// el rango ni se clasifica como cierre del bloque.
    /// </summary>
    [Fact]
    public void H1_1_ElCuadroResumenQuedaFueraDeTodoBloqueEnLasTresHojas()
    {
        AssertResumenFuera(DelimitadorBloquesHoja.R1);
        AssertResumenFuera(DelimitadorBloquesHoja.R2);
        AssertResumenFuera(DelimitadorBloquesHoja.R4);
    }

    /// <summary>
    /// H1.2/H1.3 (S2.2): geometría desplazada. Workbook SINTÉTICO con filas extra al inicio y
    /// bloques de longitud distinta (el equivalente v2→v3 del plan): el delimitador resuelve sin
    /// conocer ninguna fila. Cero cambios de código.
    /// </summary>
    [Fact]
    public void H1_1_GeometriaDesplazada_ResuelveSinCambiarCodigo()
    {
        using var libro = LibroSintetico.CrearR1(desplazamientoInicial: 3, filasExtraPorBloque: [0, 4, 1, 7, 2]);
        using var documento = SpreadsheetDocument.Open(libro.Ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");

        var bloques = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1);

        Assert.Equal(5, bloques.Count);
        for (var i = 0; i < bloques.Count; i++)
        {
            Assert.Equal(i + 1, bloques[i].AseId);
            Assert.Equal(libro.Cabeceras[i], bloques[i].Cabecera);
            Assert.Equal(libro.GranTotales[i], bloques[i].GranTotal);
            Assert.Equal(libro.EncabezadosColumnas[i], bloques[i].EncabezadoColumnas);
            Assert.Equal(2, bloques[i].Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion));
            Assert.Equal(1, bloques[i].Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSubseccion));

            // Las filas extra son filas de DATO: se clasifican por firma, no por posición.
            var datos = bloques[i].Cuenta(DelimitadorBloquesHoja.TipoFilaBloqueHoja.Dato);
            Assert.Equal(1 + libro.FilasExtraPorBloque[i], datos);

            // El cuadro resumen sintético (con la firma trampa) queda fuera del bloque.
            Assert.True(
                bloques[i].GranTotal < libro.Resumen[i],
                $"El resumen sintético (fila {libro.Resumen[i]}) no puede quedar dentro del bloque del ASE {bloques[i].AseId} (gran total {bloques[i].GranTotal}).");
        }
    }

    /// <summary>
    /// H1.1 (fail-fast accionable, G-TRAZA): si la hoja no declara el bloque de un ASE del catálogo,
    /// el error nombra hoja + ASE + rótulo esperado y NO remapea por posición.
    /// </summary>
    [Fact]
    public void H1_1_BloqueAusente_FailFastAccionablePorRotulo()
    {
        using var libro = LibroSintetico.CrearR1(desplazamientoInicial: 3, filasExtraPorBloque: [0, 1, 2, 1, 0], omitirAse: 3);
        using var documento = SpreadsheetDocument.Open(libro.Ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");

        var ex = Assert.Throws<CalculoInvalidoException>(
            () => DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains(DelimitadorBloquesHoja.NombreHojaR1, ex.Message, StringComparison.Ordinal);
        Assert.Contains("ASE 3", ex.Message, StringComparison.Ordinal);
        Assert.Contains(AseFactory.DesdeId(3).NombreCompleto, ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// H1.2/H1.4 (CA-B1, parcial H1): el escritor de slots R1 procesa los 5 bloques de la v3 con las
    /// fuentes reales de julio‑2 SIN <c>ERR-SIN-SLOT</c> por ventana, y el cierre de sección
    /// <c>Subs/Cont</c> —el que en la v2 vivía dentro de la ventana congelada y en la v3 quedó fuera—
    /// recibe el subtotal que la fuente declara. Es el defecto de §0.4-f reproducido y cerrado.
    /// </summary>
    [Fact]
    public void H1_2_R1_EscribeLosCincoBloquesSobreLaV3_SinErrSinSlotPorVentana()
    {
        var fuentes = Enumerable.Range(1, 5)
            .Select(ase => new ExcelDataReaderWorkbookLeafInputReader()
                .LeerEspejoR1(AseFactory.DesdeId(ase), FuenteR1Julio2(ase)))
            .ToList();

        using var copia = new CopiaDePlantillaV3();
        using (var documento = SpreadsheetDocument.Open(copia.Ruta, true))
        {
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            foreach (var fuente in fuentes)
            {
                var resultado = EscritorValoresSlotR1.Escribir(parte, fuente);
                Assert.True(
                    resultado.FilasColocadas > 0,
                    $"El bloque del ASE {fuente.Ase.Id} de julio-2 no colocó ninguna fila (¿bloque mal delimitado?).");
            }

            parte.Workbook?.Save();
        }

        using (var lectura = SpreadsheetDocument.Open(copia.Ruta, false))
        {
            var parte = lectura.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var bloques = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1);
            Assert.Equal(5, bloques.Count);

            for (var i = 0; i < bloques.Count; i++)
            {
                var cierrePlantilla = bloques[i].Cierres(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion)[^1];
                var cierreFuente = fuentes[i].Filas.Last(f => f.EsSubsContTotal);
                Assert.Equal(
                    cierreFuente.Valor("Total"),
                    TestHelpers.LeerCeldaNumerica(copia.Ruta, DelimitadorBloquesHoja.NombreHojaR1, $"F{cierrePlantilla.Fila}"));
            }
        }
    }

    // ── Aserciones compartidas ──────────────────────────────────────────────────────────────────

    private static IReadOnlyList<DelimitadorBloquesHoja.BloqueHoja> Delimitar(DelimitadorBloquesHoja.HojaDelimitable hoja)
    {
        using var documento = SpreadsheetDocument.Open(PlantillaV3, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        return DelimitadorBloquesHoja.Delimitar(parte, hoja);
    }

    private static void AssertGeometria(
        IReadOnlyList<DelimitadorBloquesHoja.BloqueHoja> bloques,
        (int Ase, int Cabecera, int GranTotal)[] esperado,
        (int Ase, int InicioRegion, int FinRegion)[] regiones,
        DelimitadorBloquesHoja.HojaDelimitable hoja)
    {
        Assert.Equal(esperado.Length, bloques.Count);
        for (var i = 0; i < esperado.Length; i++)
        {
            Assert.Equal(esperado[i].Ase, bloques[i].AseId);
            Assert.Equal(esperado[i].Cabecera, bloques[i].Cabecera);
            Assert.Equal(esperado[i].GranTotal, bloques[i].GranTotal);
            Assert.Equal(regiones[i].InicioRegion, bloques[i].Cabecera);
            Assert.True(
                bloques[i].GranTotal < regiones[i].FinRegion,
                $"{hoja.Nombre}: el bloque del ASE {bloques[i].AseId} debe cortar en su gran total ({bloques[i].GranTotal}) ANTES del fin de la región medida en H0 ({regiones[i].FinRegion}); el hueco contiene el cuadro resumen.");

            // Cabecera del bloque siguiente = fin de región + 1 (dato del T0 §1.3).
            if (i + 1 < esperado.Length)
            {
                Assert.Equal(regiones[i].FinRegion + 1, esperado[i + 1].Cabecera);
            }

            // La fila de encabezados de columna y la primera fila de datos siguen el contrato
            // cabecera+1 / cabecera+2 medido en el T0 — aserción de TEST, nunca constante de producción.
            Assert.Equal(bloques[i].Cabecera + 1, bloques[i].EncabezadoColumnas);
            Assert.Equal(bloques[i].Cabecera + 2, bloques[i].FilasSlot[0].Fila);
            Assert.Equal(bloques[i].GranTotal, bloques[i].FilasSlot[^1].Fila);
        }
    }

    /// <summary>
    /// Verifica que ninguna fila del cuadro resumen —las que reutilizan las firmas de cierre
    /// (<c>OPORTUNO</c>/<c>EXTEMPORANEO</c> junto a la marca <c>Total</c>)— pertenece a un bloque.
    /// El hueco donde vive el resumen es (gran total del bloque, cabecera del siguiente) y el assert
    /// exige que ese hueco CONTENGA esas firmas: si no las contuviera, el fixture sería una
    /// tautología (T0 §2.5).
    /// </summary>
    private static void AssertResumenFuera(DelimitadorBloquesHoja.HojaDelimitable hoja)
    {
        using var documento = SpreadsheetDocument.Open(PlantillaV3, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloques = DelimitadorBloquesHoja.Delimitar(parte, hoja);
        var worksheet = WorksheetDe(parte, hoja.Nombre);
        var ultimaFila = (int)worksheet.Descendants<Row>().Max(f => f.RowIndex?.Value ?? 0);

        var rangos = bloques.Select(b => (b.Cabecera, b.GranTotal)).ToList();
        var trampas = new List<int>();
        for (var i = 0; i < bloques.Count; i++)
        {
            var desde = bloques[i].GranTotal + 1;
            var hasta = i + 1 < bloques.Count ? bloques[i + 1].Cabecera : ultimaFila + 1;
            trampas.AddRange(FilasConFirmaDeResumen(parte, worksheet, desde, hasta));
        }

        Assert.NotEmpty(trampas);
        Assert.All(trampas, fila => Assert.DoesNotContain(rangos, r => fila >= r.Cabecera && fila <= r.GranTotal));
    }

    /// <summary>
    /// Filas del rango que declaran la firma del cuadro resumen (rótulo de oportunidad + marca
    /// <c>Total</c>) — el vocabulario medido en el T0 §2.5, leído de la propia hoja.
    /// </summary>
    private static IEnumerable<int> FilasConFirmaDeResumen(WorkbookPart parte, Worksheet worksheet, int desdeExclusivo, int hastaExclusivo)
    {
        foreach (var fila in worksheet.Descendants<Row>())
        {
            var numero = (int)(fila.RowIndex?.Value ?? 0);
            if (numero <= desdeExclusivo || numero >= hastaExclusivo)
            {
                continue;
            }

            var textos = new[] { "A", "B", "C", "D", "E" }
                .Select(col => LectorTextoHoja.TextoDeColumna(parte, fila, col).Trim())
                .ToArray();
            var marcaTotal = textos.Any(t => string.Equals(t, "Total", StringComparison.OrdinalIgnoreCase));
            var rotuloOportunidad = textos.Any(t =>
                string.Equals(t, "OPORTUNO", StringComparison.OrdinalIgnoreCase)
                || string.Equals(t, "EXTEMPORANEO", StringComparison.OrdinalIgnoreCase));
            if (marcaTotal && rotuloOportunidad)
            {
                yield return numero;
            }
        }
    }

    private static Worksheet WorksheetDe(WorkbookPart parte, string nombreHoja)
    {
        var hoja = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet!;
    }

    // ── Rutas explícitas (prohibido Insumos.Raiz()) ─────────────────────────────────────────────

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

    private static string PlantillaV3 => Path.Combine(
        RaizRepo(), "Docs", "Plantilla estandarizada", "v3", "Formato Remuneracion 2026xx-x Total v3.xlsx");

    /// <summary>
    /// Fuente R1 de julio‑2 resuelta por PATRÓN, tanto la carpeta del ASE (<c>N-*</c>) como el archivo
    /// (<c>Recaudoporcomponente*.xlsx</c>): nunca por nombre exacto. El commit <c>88b3acb</c> renombró
    /// los insumos y congelar el nombre es la misma clase de defecto que este plan elimina (además,
    /// la carpeta del ASE 4 en disco va sin la tilde del catálogo).
    /// </summary>
    private static string FuenteR1Julio2(int ase)
    {
        var raizInsumos = Path.Combine(RaizRepo(), "Docs", "Prueba Julio-2", "Insumos");
        var carpetas = Directory.EnumerateDirectories(raizInsumos, $"{ase}-*", SearchOption.TopDirectoryOnly)
            .OrderBy(f => f, StringComparer.Ordinal)
            .ToList();
        Assert.Single(carpetas);

        var archivos = Directory.EnumerateFiles(carpetas[0], "Recaudoporcomponente*.xlsx", SearchOption.TopDirectoryOnly)
            .OrderBy(f => f, StringComparer.Ordinal)
            .ToList();
        Assert.Single(archivos);
        return archivos[0];
    }

    /// <summary>Copia temporal de la plantilla v3 (el original nunca se abre en edición).</summary>
    private sealed class CopiaDePlantillaV3 : IDisposable
    {
        private readonly string _directorio;

        public CopiaDePlantillaV3()
        {
            _directorio = Path.Combine(Path.GetTempPath(), "p59-h1-tpl-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directorio);
            Ruta = Path.Combine(_directorio, "salida.xlsx");
            File.Copy(PlantillaV3, Ruta, overwrite: false);
        }

        public string Ruta { get; }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_directorio))
                {
                    Directory.Delete(_directorio, recursive: true);
                }
            }
            catch
            {
                // best-effort
            }
        }
    }

    // ── Workbook sintético de geometría DESPLAZADA (sin tocar la plantilla real) ────────────────

    /// <summary>
    /// Construye un workbook mínimo con la MISMA gramática de R1 (cabecera de bloque con el rótulo
    /// del ASE en B, fila de encabezados de columna, filas de dato, cierres de canal/empresa/
    /// subsección/sección y gran total) más un cuadro resumen con la firma trampa. Todos los rótulos
    /// salen del catálogo de dominio: cero nombres de empresa escritos a mano.
    /// </summary>
    private sealed class LibroSintetico : IDisposable
    {
        private readonly string _directorio;

        private LibroSintetico(string directorio, string ruta)
        {
            _directorio = directorio;
            Ruta = ruta;
        }

        public string Ruta { get; }

        public List<int> Cabeceras { get; } = [];

        public List<int> GranTotales { get; } = [];

        public List<int> EncabezadosColumnas { get; } = [];

        public List<int> Resumen { get; } = [];

        public List<int> FilasExtraPorBloque { get; } = [];

        public static LibroSintetico CrearR1(int desplazamientoInicial, int[] filasExtraPorBloque, int? omitirAse = null)
        {
            var directorio = Path.Combine(Path.GetTempPath(), "p59-h1-" + Guid.NewGuid().ToString("N"));
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

                var fila = 1;
                for (var i = 0; i < desplazamientoInicial; i++)
                {
                    Fila(datos, fila++, []);
                }

                for (var ase = 1; ase <= 5; ase++)
                {
                    var extra = filasExtraPorBloque[ase - 1];
                    libro.FilasExtraPorBloque.Add(extra);
                    var omitido = omitirAse == ase;
                    var rotulo = AseFactory.DesdeId(ase).NombreCompleto.ToUpperInvariant();

                    libro.Cabeceras.Add(fila);
                    Fila(datos, fila++, [("B", omitido ? "SIN ROTULO" : rotulo)]);
                    libro.EncabezadosColumnas.Add(fila);
                    Fila(datos, fila++, [("F", "Total"), ("G", "Componente TDF")]);

                    for (var k = 0; k < extra; k++)
                    {
                        FilaNumerica(datos, fila++, ("E", "Vlr Servicio"));
                    }

                    FilaNumerica(datos, fila++, ("E", "Vlr Servicio"));
                    FilaNumerica(datos, fila++, ("D", "1"), ("E", "Total"));
                    FilaNumerica(datos, fila++, ("C", "EMPRESA UNO"), ("D", "Total"));
                    FilaNumerica(datos, fila++, ("B", "SUBSECCION UNO"), ("C", "Total"));
                    FilaNumerica(datos, fila++, ("A", "SECCION UNO"), ("B", "Total"));
                    FilaNumerica(datos, fila++, ("A", "SECCION DOS"), ("B", "Total"));
                    libro.GranTotales.Add(fila);
                    FilaNumerica(datos, fila++, ("A", "Total"));

                    // Cuadro resumen sintético: reutiliza la firma de cierre de canal y de sección.
                    libro.Resumen.Add(fila);
                    FilaNumerica(datos, fila++, ("C", "OPORTUNO"), ("D", "Total"));
                    FilaNumerica(datos, fila++, ("A", "EMPRESA UNO"), ("B", "OPORTUNO"), ("C", "Total"));
                    Fila(datos, fila++, []);
                }
            }

            return libro;
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_directorio))
                {
                    Directory.Delete(_directorio, recursive: true);
                }
            }
            catch
            {
                // best-effort
            }
        }

        private static void Fila(SheetData datos, int numero, (string Columna, string Texto)[] textos)
        {
            var fila = new Row { RowIndex = (uint)numero };
            foreach (var (columna, texto) in textos)
            {
                fila.Append(new Cell
                {
                    CellReference = $"{columna}{numero}",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text(texto)),
                });
            }

            datos.Append(fila);
        }

        private static void FilaNumerica(SheetData datos, int numero, params (string Columna, string Texto)[] textos)
        {
            var fila = new Row { RowIndex = (uint)numero };
            foreach (var (columna, texto) in textos)
            {
                fila.Append(new Cell
                {
                    CellReference = $"{columna}{numero}",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text(texto)),
                });
            }

            // Valor numérico en la zona de componentes: la fila NO es un encabezado de columna.
            fila.Append(new Cell
            {
                CellReference = $"F{numero}",
                DataType = CellValues.Number,
                CellValue = new CellValue("0"),
            });
            datos.Append(fila);
        }
    }
}
