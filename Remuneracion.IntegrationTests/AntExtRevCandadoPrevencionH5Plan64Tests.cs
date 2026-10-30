using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 64 (H5.1/H5.2/H5.3 — REQ-10): PREVENCIÓN SISTÉMICA del frente ANT EXT-REV.
///
/// <para><b>H5.1 — candado anti-quemado extendido a los caminos tocados.</b> Dos capas, como los
/// candados de los Planes 55/56/59:
/// <list type="bullet">
/// <item><b>ESTÁTICA:</b> en los caminos que tocan el insumo y la hoja (modelo, lector, resolvedor,
/// región del writer, preflight y finder) no puede haber literales de celda/rango/fila/letra ni ramas
/// por ASE, período, quincena, empresa o fuente. La plantilla es la autoridad de los destinos y el
/// insumo aporta la forma; el código solo compara DATOS. Las excepciones son una lista CERRADA con
/// motivo declarado (no un liberalismo).</item>
/// <item><b>INJERTO:</b> cada forma prohibida inyectada a propósito debe ENROJECER el analizador, incluido
/// el injerto sobre una COPIA temporal de un archivo real (que se revierte al borrarla): un candado que
/// no se puede probar en rojo es decorativo.</item>
/// <item><b>CONDUCTUAL:</b> el MISMO insumo reordenado (grupos y conceptos) produce la MISMA colocación
/// celda a celda, y una plantilla con la geometría DESPLAZADA (otras filas y otras columnas) se resuelve
/// igual por etiquetas: la invariante es la gramática, no la coordenada.</item>
/// </list></para>
///
/// <para><b>H5.2 — sonda de etiquetas del insumo.</b> Una etiqueta que el catálogo no tabula deja un
/// hallazgo durable con forma administrativa <c>ASE n · RecaudosReversados · Sheet1 · {etiqueta}</c>; si
/// además COEXISTE con una fuente del catálogo ausente, el lector hace fail-fast <c>ERR-FORMATO-FUENTE</c>
/// (gemelo A3 de los Planes 50/51). Se cierra como candado VERDE: la sonda ya vive en el lector de H1 y
/// en la propagación de H3.</para>
///
/// <para><b>H5.3 — doctrina de fixtures.</b> Prohibido inventar valores del insumo: los asserts se anclan
/// a la tabla medida por H0 (<c>Docs/T0-Plan64-Ant-Ext-Rev.md</c>) o al manual del administrativo, y los
/// montos SINTÉTICOS se declaran como tales (solo forma, nunca un monto real del negocio). Prohibidos los
/// asserts tautológicos — la identidad de una fórmula propia no prueba nada (precedente Plan 61 H5.4).</para>
/// </summary>
public sealed class AntExtRevCandadoPrevencionH5Plan64Tests
{
    private const decimal Tolerancia = 0.5m;

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Rutas explícitas (prohibido Insumos.Raiz() en lo nuevo, §2.1)
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

    private static string RutaDeTests => Ruta("Remuneracion.IntegrationTests");

    private static string PlantillaV4() =>
        Ruta("Docs", "Plantilla estandarizada", "v4", "Formato Remuneracion 2026xx-x Total v4.xlsx");

    private static readonly (string Periodo, string Carpeta)[] Periodos =
    [
        ("2026071", "Prueba Julio-1"),
        ("2026072", "Prueba Julio-2"),
        ("2026082", "Prueba Agosto-2"),
        ("2026091", "Prueba Septiembre 1"),
        ("2026092", "Prueba Septiembre-2"),
        ("2026101", "Prueba Octubre-1")
    ];

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia"
    ];

    private static string Insumo(string periodo, int aseId)
    {
        var carpeta = Path.Combine(
            Ruta("Docs", Periodos.Single(p => p.Periodo == periodo).Carpeta, "Insumos"),
            CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta del ASE{aseId} en {periodo}: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .SingleOrDefault(f => Path.GetFileName(f).StartsWith("RecaudosReversados_", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta un único RecaudosReversados_*.xlsx en {carpeta}");
        return archivo!;
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H5.1 — candado ESTÁTICO
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>Archivos completos del camino nuevo de Plan 64.</summary>
    private static readonly string[] CaminosCompletos =
    [
        Path.Combine("Remuneracion.Core", "Models", "AntExtRevAseInputs.cs"),
        Path.Combine("Remuneracion.Core", "Models", "CatalogoFuentesAntExtRev.cs"),
        Path.Combine("Remuneracion.Infrastructure", "Excel", "ExcelDataReaderWorkbookLeafInputReader.AntExtRev.cs"),
        Path.Combine("Remuneracion.Infrastructure", "Excel", "ResolvedorDestinoAntExtRev.cs")
    ];

    /// <summary>Regiones acotadas de archivos compartidos que el Plan 64 tocó (no se juzga el resto).</summary>
    private static readonly (string Archivo, string Desde, string Hasta)[] Regiones =
    [
        (
            Path.Combine("Remuneracion.Infrastructure", "Excel", "OpenXmlPlantillaWriter.cs"),
            "internal sealed record CeldaAntExtRev(",
            "private static void EscribirCeldasDetalleR2R4("),
        (
            Path.Combine("Remuneracion.Core", "Services", "ProcesadorPeriodo.cs"),
            "var rutaRecaudosReversados = _localizador.BuscarRecaudosReversados(carpetaAse)",
            "leaf.AntExtRev = _leafReader.LeerAntExtRev(ase, solicitud.Periodo, rutaRecaudosReversados);"),
        (
            Path.Combine("Remuneracion.Core", "Services", "ValidadorInsumosPeriodo.cs"),
            "var rutaRecaudosReversados = _localizador.BuscarRecaudosReversados(carpetaAse)",
            "AgregarSiFirmaInvalida(rutaRecaudosReversados, $\"ASE {aseId}\""),
        (
            Path.Combine("Remuneracion.Core", "Services", "FormateadorInsumosFaltantes.cs"),
            "ReporteInsumoAse.RecaudosReversados =>",
            "_ => throw new ArgumentOutOfRangeException(nameof(reporte)"),
        (
            Path.Combine("Remuneracion.Infrastructure", "FileSystem", "ArchivoFuenteLocator.cs"),
            "public string? BuscarRecaudosReversados(string carpetaAse) =>",
            "BuscarArchivoNormalizado(carpetaAse, \"recaudosreversados\");")
    ];

    /// <summary>Rótulos del catálogo de fuentes que el camino no puede escribir como literal.</summary>
    private static readonly string[] RotulosDeFuente = ["OCCIDENTE", "NUEVO ESQUEMA", "ENEL", "ENERBIT", "CIUDAD LIMPIA"];

    /// <summary>
    /// Lista CERRADA de tokens que NO son coordenadas: (archivo, token, motivo). Un token solo queda
    /// exento si el archivo coincide Y el literal lo contiene tal cual.
    /// </summary>
    private static readonly (string Archivo, string Token, string Motivo)[] Exentos =
    [
        (
            Path.Combine("Remuneracion.Core", "Models", "CatalogoFuentesAntExtRev.cs"),
            "F199",
            "CITA de la sinonimia del Plan 34 (oráculo Reporte Componentes R1!F199); vive en TEXT de la tabla "
            + "citada de fuentes, no es una coordenada de destino de ANT EXT-REV."),
        (
            Path.Combine("Remuneracion.Core", "Models", "CatalogoFuentesAntExtRev.cs"),
            "R1",
            "CITA del oráculo de la sinonimia (el REPORTE R1), no una celda: la tabla citada es DATO con su cita."),
        (
            Path.Combine("Remuneracion.Core", "Models", "CatalogoFuentesAntExtRev.cs"),
            "T0",
            "CITA del documento de evidencia (Docs/T0-Plan64-Ant-Ext-Rev.md) dentro del texto de la tabla citada."),
        (
            Path.Combine("Remuneracion.Core", "Services", "ProcesadorPeriodo.cs"),
            "R1",
            "NOMBRE administrativo del reporte en los mensajes de progreso preexistentes (\"leyendo R1, R2 y R4\"), "
            + "no una coordenada ni una rama."),
        (
            Path.Combine("Remuneracion.Core", "Services", "ProcesadorPeriodo.cs"),
            "R2",
            "NOMBRE administrativo del reporte en los mensajes de progreso preexistentes, no una coordenada."),
        (
            Path.Combine("Remuneracion.Core", "Services", "ProcesadorPeriodo.cs"),
            "R4",
            "NOMBRE administrativo del reporte en los mensajes de progreso preexistentes, no una coordenada."),
        (
            Path.Combine("Remuneracion.Infrastructure", "Excel", "OpenXmlPlantillaWriter.cs"),
            "T0",
            "CITA del documento de evidencia dentro del mensaje del catálogo de conceptos; no selecciona ninguna celda."),
        (
            Path.Combine("Remuneracion.Infrastructure", "Excel", "ResolvedorDestinoAntExtRev.cs"),
            "B",
            "GRAMÁTICA declarada de la hoja: la columna donde viven las etiquetas de bloque. Es la única "
            + "propiedad que no se puede derivar del contenido (todas las etiquetas comparten columna) y se "
            + "declara UNA vez con cita, como el resto de la solución con la columna de rótulos de un bloque.")
    ];

    /// <summary>Celda o rango A1 dentro de un literal (<c>C11</c>, <c>C11:L12</c>, <c>$D$9</c>).</summary>
    private static readonly Regex ReferenciaDeCelda = new(
        @"(?<![A-Za-z0-9_])\$?[A-Z]{1,2}\$?[0-9]{1,4}(?![0-9])",
        RegexOptions.Compiled);

    /// <summary>Letra de columna suelta (<c>"B"</c>, <c>"AB"</c>).</summary>
    private static readonly Regex LetraDeColumnaSuelta = new(@"^[A-Z]{1,2}$", RegexOptions.Compiled);

    [Fact]
    public void H5_1_CaminosTocados_SinCoordenadasNiRamasQuemadas()
    {
        var violaciones = new List<string>();

        foreach (var relativo in CaminosCompletos)
        {
            var ruta = Ruta(relativo);
            Assert.True(File.Exists(ruta), $"Falta el archivo del camino tocado: {relativo}");
            violaciones.AddRange(Analizar(File.ReadAllText(ruta), relativo));
        }

        foreach (var (archivo, desde, hasta) in Regiones)
        {
            var ruta = Ruta(archivo);
            Assert.True(File.Exists(ruta), $"Falta el archivo del camino tocado: {archivo}");
            var region = TextoDeRegion(File.ReadAllText(ruta), desde, hasta, archivo);
            violaciones.AddRange(Analizar(region, $"{archivo} [{desde}..]"));
        }

        Assert.True(
            violaciones.Count == 0,
            "G-QUEMADO/REQ-10: hay coordenadas o ramas quemadas en los caminos de ANT EXT-REV "
            + "(la plantilla es la autoridad de los destinos; el insumo aporta la forma):"
            + Environment.NewLine
            + string.Join(Environment.NewLine, violaciones));
    }

    [Fact]
    public void H5_1_Injerto_CoordenadasDeCeldaYRango_Enrojecen()
    {
        var archivo = Path.Combine("Remuneracion.Infrastructure", "Excel", "ExcelDataReaderWorkbookLeafInputReader.AntExtRev.cs");

        Assert.NotEmpty(Analizar("""var destino = "C11";""", archivo));
        Assert.NotEmpty(Analizar("""var rango = "C11:L12";""", archivo));
        Assert.NotEmpty(Analizar("""var celda = "$D$9";""", archivo));
        Assert.NotEmpty(Analizar("""throw new Exception("la celda F43 no existe");""", archivo));
    }

    [Fact]
    public void H5_1_Injerto_LetraYFilaQuemadas_Enrojecen()
    {
        var archivo = Path.Combine("Remuneracion.Infrastructure", "Excel", "ResolvedorDestinoAntExtRev.cs");

        // La letra pasa SOLO por la lista cerrada con motivo; otra letra en el resolvedor es un quema.
        Assert.NotEmpty(Analizar("""private const string ColumnaConcepto = "C";""", archivo));
        Assert.NotEmpty(Analizar("""var filaAplicado = 11;""", archivo));
        Assert.NotEmpty(Analizar("""fila.RowIndex = 40;""", archivo));
    }

    [Fact]
    public void H5_1_Injerto_RamasPorAsePeriodoQuincena_Enrojecen()
    {
        var archivo = Path.Combine("Remuneracion.Infrastructure", "Excel", "ResolvedorDestinoAntExtRev.cs");

        Assert.NotEmpty(Analizar("""var destino = aseId switch { 1 => "C11", 4 => "D11", _ => "E11" };""", archivo));
        Assert.NotEmpty(Analizar("""if (ase.Id == 4) { return; }""", archivo));
        Assert.NotEmpty(Analizar("""if (periodo.NumeroQuincena == 2) { return; }""", archivo));
        Assert.NotEmpty(Analizar("""if (periodo.CodigoCompleto == "2026101") { return; }""", archivo));
    }

    [Fact]
    public void H5_1_Injerto_LiteralDeFuente_Enrojece()
    {
        var archivo = Path.Combine("Remuneracion.Infrastructure", "Excel", "ExcelDataReaderWorkbookLeafInputReader.AntExtRev.cs");

        Assert.NotEmpty(Analizar("""var fuente = "OCCIDENTE";""", archivo));
        Assert.NotEmpty(Analizar("""if (etiqueta == "NUEVO ESQUEMA") { return; }""", archivo));
        Assert.NotEmpty(Analizar("""var bloque = "ENEL";""", archivo));
    }

    /// <summary>
    /// INJERTO sobre una COPIA temporal de un archivo real: el analizador que juzga el camino productivo
    /// detecta la forma inyectada y, al borrar la copia, el repositorio queda intacto (el injerto se
    /// REVIERTE; nunca se edita el archivo real).
    /// </summary>
    [Fact]
    public void H5_1_Injerto_EnCopiaTemporal_EnrojeceYSeRevierte()
    {
        var relativo = Path.Combine("Remuneracion.Infrastructure", "Excel", "ResolvedorDestinoAntExtRev.cs");
        var original = File.ReadAllText(Ruta(relativo));
        var copia = Path.Combine(Path.GetTempPath(), $"p64-h5-injerto-{Guid.NewGuid():N}.cs");

        try
        {
            File.WriteAllText(copia, original.Replace(
                "internal static class ResolvedorDestinoAntExtRev",
                "internal static class ResolvedorDestinoAntExtRev // injerto\n    internal const string CeldaQuemada = \"K84\";",
                StringComparison.Ordinal));

            var violaciones = Analizar(File.ReadAllText(copia), relativo);
            Assert.Contains(violaciones, v => v.Contains("K84", StringComparison.Ordinal));

            // La copia limpia (el original) NO tiene violaciones: el rojo es del injerto, no del camino.
            Assert.Empty(Analizar(original, relativo));
        }
        finally
        {
            File.Delete(copia);
        }

        Assert.Equal(original, File.ReadAllText(Ruta(relativo)));
    }

    [Fact]
    public void H5_1_ElAnalizador_NoEnrojeceConLaMaquinariaLegitima()
    {
        // Allow-list viva: comparar DOS datos del dominio, iterar por identidad, resolver por catálogo o
        // leer etiquetas de la plantilla NO es quemar nada.
        const string legitimo = """
            foreach (var destino in destinos.Where(d => d.AseId == leaf.Ase.Id)) { }
            var fila = CatalogoFuentesAntExtRev.BuscarPorEtiquetaBloque(etiquetaBloque);
            var mismo = SinonimosEmpresaR1.SonMismaEmpresa(origen, candidato);
            var texto = LectorTextoHoja.TextoDeColumna(parte, fila, ColumnaEtiquetaBloque);
            var ase = AseFactory.DesdeId(aseId);
            var referencia = $"{columna}{fila.RowIndex!.Value.ToString(CultureInfo.InvariantCulture)}";
            """;

        Assert.Empty(Analizar(legitimo, Path.Combine("Remuneracion.Infrastructure", "Excel", "ResolvedorDestinoAntExtRev.cs")));

        // La tabla de etiquetas de bloque es DATO citado: sus literales viven en el catálogo, no se dispersan.
        Assert.Empty(Analizar(
            """var etiqueta = "BANCO DE OCCIDENTE";""",
            Path.Combine("Remuneracion.Core", "Models", "CatalogoFuentesAntExtRev.cs")));
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H5.1 — candado CONDUCTUAL
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// El MISMO insumo con los grupos de fuente y los conceptos en OTRO ORDEN produce la MISMA colocación
    /// celda a celda: el destino lo resuelve la identidad (fuente × concepto × ASE), nunca la posición.
    /// </summary>
    [Fact]
    public void H5_1_Conductual_MismoInsumoReordenado_ColocaIdentico()
    {
        var insumoA = CrearInsumoSintetico(
        [
            Grupo("OCCIDENTE", 100m, 10m, -5m, 0m),
            Grupo("NUEVO ESQUEMA", 200m, 20m, -7m, 0m),
            Grupo("ENEL", 300m, 30m, -9m, 0m)
        ]);
        var insumoB = CrearInsumoSintetico(
        [
            Grupo("ENEL", 300m, 30m, -9m, 0m, invertirConceptos: true),
            Grupo("OCCIDENTE", 100m, 10m, -5m, 0m, invertirConceptos: true),
            Grupo("NUEVO ESQUEMA", 200m, 20m, -7m, 0m, invertirConceptos: true)
        ]);

        var plantilla = CopiarPlantilla();
        try
        {
            var leidoA = LeerSintetico(insumoA, aseId: 4);
            var leidoB = LeerSintetico(insumoB, aseId: 4);

            Assert.Equal(NormalizarFuentes(leidoA), NormalizarFuentes(leidoB));
            Assert.Equal(PlanDeColocacion(plantilla, leidoA), PlanDeColocacion(plantilla, leidoB));

            // Tres fuentes presentes × cuatro conceptos = 12 celdas escritas, sin importar el orden de entrada.
            Assert.Equal(12, PlanDeColocacion(plantilla, leidoA).Count(entrada => entrada.SeEscribe));
        }
        finally
        {
            File.Delete(insumoA);
            File.Delete(insumoB);
            File.Delete(plantilla);
        }
    }

    /// <summary>
    /// Una plantilla con la geometría DESPLAZADA (bloques en otras filas y las columnas de valor en
    /// <c>D..M</c> en vez de <c>C..L</c>) se resuelve igual por etiquetas: la identidad manda, no la
    /// coordenada congelada de la v4.
    /// </summary>
    [Fact]
    public void H5_1_Conductual_GeometriaDesplazada_ResuelvePorEtiquetas()
    {
        var plantilla = CrearPlantillaDesplazada();
        try
        {
            using var documento = SpreadsheetDocument.Open(plantilla, false);
            var destinos = ResolvedorDestinoAntExtRev.Resolver(documento.WorkbookPart!);

            // 3 bloques del catálogo × 5 ASE × 2 reportes × 2 filas de valor = 60 celdas.
            Assert.Equal(60, destinos.Count);

            var anticiposEnelAse4 = ResolvedorDestinoAntExtRev.Buscar(destinos, "ENEL", 4, "ANTICIPOS", "REPORTE 4");
            Assert.NotNull(anticiposEnelAse4);
            Assert.Equal("K84", anticiposEnelAse4!.Celda);

            var aplicadoOccidenteAse1 = ResolvedorDestinoAntExtRev.Buscar(destinos, "BANCO DE OCCIDENTE", 1, "APLICADO", "REPORTE 3");
            Assert.NotNull(aplicadoOccidenteAse1);
            Assert.Equal("D43", aplicadoOccidenteAse1!.Celda);

            // Todas las celdas viven en la geometría desplazada: ninguna coordenada congelada de la v4.
            Assert.All(destinos, destino =>
            {
                Assert.Matches("^[D-M](43|44|63|64|83|84)$", destino.Celda);
                Assert.NotEqual("C11", destino.Celda);
            });
        }
        finally
        {
            File.Delete(plantilla);
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H5.2 — Sonda de etiquetas del insumo
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Los 30 insumos reales no tienen NINGUNA etiqueta sin mapear: la sonda está verde sobre el corpus
    /// medido por H0 (si un período trae una fuente nueva, este test la nombra antes de que el dinero se
    /// pierda).
    /// </summary>
    [Fact]
    public void H5_2_Sonda_Los30InsumosReales_NoTienenEtiquetasSinMapear()
    {
        foreach (var (periodo, _) in Periodos)
        {
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var leido = Leer(Insumo(periodo, aseId), aseId);

                Assert.True(
                    leido.HallazgosEtiquetasNoMapeadas.Count == 0,
                    $"Sonda H5.2 · {periodo} ASE{aseId}: {string.Join(" | ", leido.HallazgosEtiquetasNoMapeadas)}");
                Assert.Empty(leido.EtiquetasNoReconocidas);
            }
        }
    }

    /// <summary>
    /// Una etiqueta que el catálogo NO tabula deja un hallazgo durable con la forma administrativa
    /// <c>ASE n · RecaudosReversados · Sheet1 · {etiqueta}</c>. Sin una fuente ausente en el mismo archivo
    /// es SOLO evidencia: no bloquea (gemelo A3, Plan 50/51).
    /// </summary>
    [Fact]
    public void H5_2_Sonda_EtiquetaSinMapear_HallazgoConFormaAdministrativa()
    {
        var ruta = CrearInsumoSintetico(
        [
            Grupo("OCCIDENTE", 100m),
            Grupo("NUEVO ESQUEMA", 200m),
            Grupo("ENEL", 300m),
            Grupo("EMPRESA FANTASMA", 400m)
        ]);

        try
        {
            var leido = LeerSintetico(ruta, aseId: 1);

            Assert.Equal(
                ["ASE 1 · RecaudosReversados · Sheet1 · EMPRESA FANTASMA"],
                leido.HallazgosEtiquetasNoMapeadas.ToArray());

            // La evidencia del Plan 50/51 se conserva (contrato de H1) y el lector no bloquea.
            Assert.Equal(["ANTEXTREV:EMPRESA FANTASMA"], leido.EtiquetasNoReconocidas.ToArray());
            Assert.Empty(leido.FuentesAusentes);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    /// <summary>
    /// La sonda COEXISTIENDO con una ausencia del catálogo es fail-fast (precedente Plan 50 A3): la
    /// etiqueta desconocida podría ser la fuente ausente con otro nombre y jamás se asume 0 ni se mapea por
    /// parecido.
    /// </summary>
    [Fact]
    public void H5_2_Sonda_EtiquetaSinMapearConAusencia_FailFast()
    {
        var ruta = CrearInsumoSintetico(
        [
            Grupo("ENEL", 300m),
            Grupo("EMPRESA FANTASMA", 200m)
        ]);

        try
        {
            var excepcion = Assert.Throws<CalculoInvalidoException>(() => LeerSintetico(ruta, aseId: 1));

            Assert.Equal(CodigoError.FormatoFuente, excepcion.Codigo);
            Assert.Contains("ASE 1", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("RecaudosReversados", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("Sheet1", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("EMPRESA FANTASMA", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("ANTEXTREV:OCCIDENTE", excepcion.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H5.3 — Doctrina de fixtures (candado textual)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H5_3_Candado_DoctrinaDeFixtures()
    {
        var fixtures = FixturesPlan64();
        Assert.Equal(5, fixtures.Count);

        var violaciones = new List<string>();
        foreach (var ruta in fixtures)
        {
            violaciones.AddRange(AnalizarDoctrina(File.ReadAllText(ruta), Path.GetFileName(ruta)));
        }

        Assert.True(
            violaciones.Count == 0,
            "H5.3/REQ-10: los fixtures de Plan 64 violan la doctrina (valores inventados, atajos o tautologías):"
            + Environment.NewLine
            + string.Join(Environment.NewLine, violaciones));
    }

    /// <summary>
    /// Anclas de la doctrina que se pueden verificar textualmente:
    /// (a) el oráculo medido se declara (T0 congelado en los 30 insumos y manual del administrativo en la
    /// escritura/comparador);
    /// (b) los insumos sintéticos se declaran como tales y nunca como monto real del negocio;
    /// (c) ningún assert compara un agregado contra la fórmula propia del modelo (<c>SumaConceptos</c>,
    /// <c>TotalBloque2</c>) — sería una tautología;
    /// (d) ningún test se apaga (<c>Skip</c>) ni se rellena con un assert vacío.
    /// </summary>
    [Fact]
    public void H5_3_ElCandadoDeDoctrina_EnrojeceConValoresInventadosYAtajos()
    {
        var archivo = "AntExtRevLecturaH1Plan64Tests.cs";

        Assert.NotEmpty(AnalizarDoctrina("""[Fact(Skip = "pendiente")] public void X() { }""", archivo));
        Assert.NotEmpty(AnalizarDoctrina("Assert.True(true);", archivo));
        Assert.NotEmpty(AnalizarDoctrina("Assert.False(false);", archivo));
        Assert.NotEmpty(AnalizarDoctrina("Assert.Equal(125m, fuente.SumaConceptos);", archivo));
        Assert.NotEmpty(AnalizarDoctrina("Assert.Equal(-3260930m, leido.TotalBloque2);", archivo));

        // El corpus medido y el sintético declarado NO enrojecen.
        const string legitimo = """
            // Tabla CONGELADA por H0 (T0 §4.3).
            // Los casos negativos usan insumos sintéticos: solo forma, nunca un monto real del negocio.
            Assert.Equal(-3260930m, fuente.TotalFuente);
            Assert.NotNull(leido.TotalControlBloque3);
            """;
        Assert.Empty(AnalizarDoctrina(legitimo, archivo));
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Analizador del candado anti-quemado (misma función que juzga los caminos reales)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    internal static IReadOnlyList<string> Analizar(string fuente, string archivo)
    {
        var violaciones = new List<string>();
        var lineas = SinComentarios(fuente);
        var esCatalogo = archivo.EndsWith("CatalogoFuentesAntExtRev.cs", StringComparison.OrdinalIgnoreCase);

        for (var i = 0; i < lineas.Count; i++)
        {
            var linea = lineas[i];

            // (1) Coordenadas: celda / rango / letra dentro de un literal de cadena.
            foreach (Match literal in Regex.Matches(linea, "\"([^\"\\\\]*)\""))
            {
                var texto = literal.Groups[1].Value;
                foreach (Match token in ReferenciaDeCelda.Matches(texto))
                {
                    if (EsExento(archivo, token.Value))
                    {
                        continue;
                    }

                    violaciones.Add($"{archivo}:{i + 1}: referencia de celda/rango quemada '{token.Value}' → {linea.Trim()}");
                }

                if (LetraDeColumnaSuelta.IsMatch(texto) && !EsExento(archivo, texto))
                {
                    violaciones.Add($"{archivo}:{i + 1}: letra de columna quemada '{texto}' → {linea.Trim()}");
                }

                // (2) Empresa/fuente como literal (la tabla citada de fuentes es la única excepción).
                if (!esCatalogo)
                {
                    foreach (var rotulo in RotulosDeFuente)
                    {
                        if (SinonimosEmpresaR1.SonMismaEmpresa(texto, rotulo))
                        {
                            violaciones.Add($"{archivo}:{i + 1}: literal de fuente '{texto}' → {linea.Trim()}");
                        }
                    }
                }
            }

            // (3) Fila quemada en una constante o asignación de geometría.
            if (Regex.IsMatch(linea, @"\b\w*[Ff]ila\w*\s*=\s*[1-9]\d{0,3}\s*;")
                || Regex.IsMatch(linea, @"\bRowIndex\s*=\s*[1-9]\d{0,3}\s*;"))
            {
                violaciones.Add($"{archivo}:{i + 1}: número de fila quemado → {linea.Trim()}");
            }

            // (4) Rama por ASE, período o quincena.
            if (Regex.IsMatch(linea, @"\b\w*[Aa]se(Id|\.Id)\s*(==|!=)\s*[1-5]\b")
                || Regex.IsMatch(linea, @"\b\w*[Aa]seId\s+switch\b")
                || Regex.IsMatch(linea, @"\bcase\s+[1-5]\s*:")
                || Regex.IsMatch(linea, @"\[\s*\(?\s*[1-5]\s*,\s*[1-5]\s*\)?\s*\]\s*="))
            {
                violaciones.Add($"{archivo}:{i + 1}: rama/lista por ASE → {linea.Trim()}");
            }

            if (Regex.IsMatch(linea, @"\b(Periodo\.)?(CodigoCompleto|NumeroQuincena)\s*(==|!=|switch)")
                || Regex.IsMatch(linea, @"\bQ[uU]incena\s*(==|!=)\s*[12]\b")
                || Regex.IsMatch(linea, "Periodo\\.Parse\\s*\\(\\s*\""))
            {
                violaciones.Add($"{archivo}:{i + 1}: rama por período/quincena → {linea.Trim()}");
            }
        }

        return violaciones;
    }

    /// <summary>Doctrina de fixtures: valores medidos, sintéticos declarados y cero tautologías/atajos.</summary>
    internal static IReadOnlyList<string> AnalizarDoctrina(string fuente, string archivo)
    {
        var violaciones = new List<string>();
        var lineas = SinComentarios(fuente);
        var codigo = string.Join('\n', lineas);

        var sinteticos = codigo.Contains("CrearInsumo", StringComparison.Ordinal)
            || codigo.Contains("Grupo(", StringComparison.Ordinal);

        for (var i = 0; i < lineas.Count; i++)
        {
            var linea = lineas[i];

            if (linea.Contains("Skip =", StringComparison.Ordinal)
                || linea.Contains("[Fact(Skip", StringComparison.Ordinal)
                || linea.Contains("Assert.True(true)", StringComparison.Ordinal)
                || linea.Contains("Assert.False(false)", StringComparison.Ordinal))
            {
                violaciones.Add($"{archivo}:{i + 1}: atajo que apaga o vacía la prueba → {linea.Trim()}");
            }

            // Tautología: el valor esperado es la fórmula propia del modelo, no una medición independiente.
            if (Regex.IsMatch(linea, @"Assert\.Equal\([^;]*,\s*[\w\.]+\.(SumaConceptos|TotalBloque2)\s*\)"))
            {
                violaciones.Add($"{archivo}:{i + 1}: assert tautológico sobre la fórmula propia del modelo → {linea.Trim()}");
            }

            // Identidad literal: Assert.Equal(X, X).
            var igualdad = Regex.Match(linea, @"Assert\.Equal\(\s*([^,]+?)\s*,\s*([^)]+?)\s*\)");
            if (igualdad.Success
                && string.Equals(igualdad.Groups[1].Value.Trim(), igualdad.Groups[2].Value.Trim(), StringComparison.Ordinal))
            {
                violaciones.Add($"{archivo}:{i + 1}: identidad literal (Assert.Equal(X, X)) → {linea.Trim()}");
            }
        }

        // Los montos sintéticos se declaran: solo forma, nunca un monto real del negocio.
        if (sinteticos && !fuente.Contains("nunca un monto real", StringComparison.OrdinalIgnoreCase))
        {
            violaciones.Add($"{archivo}: construye insumos sintéticos sin declarar que NO son montos reales del negocio");
        }

        return violaciones;
    }

    private static bool EsExento(string archivo, string token) =>
        Exentos.Any(exento =>
            archivo.StartsWith(exento.Archivo, StringComparison.OrdinalIgnoreCase)
            && string.Equals(exento.Token, token, StringComparison.Ordinal));

    private static string TextoDeRegion(string texto, string desde, string hasta, string archivo)
    {
        var inicio = texto.IndexOf(desde, StringComparison.Ordinal);
        Assert.True(inicio >= 0, $"No se encontró la región '{desde}' en {archivo}.");
        var fin = texto.IndexOf(hasta, inicio, StringComparison.Ordinal);
        Assert.True(fin > inicio, $"No se encontró el cierre '{hasta}' en {archivo}.");
        return texto[inicio..fin];
    }

    /// <summary>Quita comentarios: el candado juzga CÓDIGO, no prosa (las citas viven en comentarios).</summary>
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

    /// <summary>
    /// Corpus de la doctrina: los fixtures de H0..H4. El candado NO se juzga a sí mismo (contiene las formas
    /// prohibidas como texto de las pruebas de injerto, igual que el candado de período del Plan 59).
    /// </summary>
    private static List<string> FixturesPlan64() =>
        Directory.EnumerateFiles(RutaDeTests, "AntExtRev*Plan64Tests.cs", SearchOption.TopDirectoryOnly)
            .Where(ruta => !string.Equals(
                Path.GetFileName(ruta),
                "AntExtRevCandadoPrevencionH5Plan64Tests.cs",
                StringComparison.OrdinalIgnoreCase))
            .OrderBy(ruta => ruta, StringComparer.Ordinal)
            .ToList();

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Lectura sintética y colocación (reusa la unidad de producción, nunca una copia de la regla)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static AntExtRevAseInputs Leer(string ruta, int aseId) =>
        new ExcelDataReaderWorkbookLeafInputReader().LeerAntExtRev(
            new Ase { Id = aseId, NombreCorto = $"ASE{aseId}" },
            Periodo.Parse("2026091"),
            ruta);

    private static AntExtRevAseInputs LeerSintetico(string ruta, int aseId) => Leer(ruta, aseId);

    private static string NormalizarFuentes(AntExtRevAseInputs leido) =>
        string.Join(
            ";",
            leido.Fuentes
                .OrderBy(fuente => fuente.FuenteEtiqueta, StringComparer.Ordinal)
                .Select(fuente =>
                    $"{fuente.FuenteEtiqueta}=" + string.Join(
                        "/",
                        new[] { fuente.AplicadoR3, fuente.AnticiposR3, fuente.AplicadoR4, fuente.AnticiposR4, fuente.TotalFuente }
                            .Select(valor => valor.ToString(CultureInfo.InvariantCulture)))));

    private sealed record Colocacion(string Celda, decimal Valor, bool SeEscribe);

    /// <summary>Plan de colocación del writer (unidad de producción), ordenado para comparar sin orden.</summary>
    private static IReadOnlyList<Colocacion> PlanDeColocacion(string plantilla, AntExtRevAseInputs leido)
    {
        var leaf = new WorkbookLeafInputs { Ase = leido.Ase, AntExtRev = leido };
        using var documento = SpreadsheetDocument.Open(plantilla, false);
        var plan = OpenXmlPlantillaWriter.PlanificarAntExtRev(documento.WorkbookPart!, [leaf]);
        Assert.NotNull(plan);
        return plan!
            .OrderBy(celda => celda.Celda, StringComparer.Ordinal)
            .Select(celda => new Colocacion(celda.Celda, celda.Valor, celda.SeEscribe))
            .ToList();
    }

    private static string CopiarPlantilla()
    {
        var destino = Path.Combine(Path.GetTempPath(), $"p64-h5-plantilla-{Guid.NewGuid():N}.xlsx");
        File.Copy(PlantillaV4(), destino, overwrite: true);
        return destino;
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Insumo sintético con la gramática MEDIDA (solo forma, nunca un monto real del negocio)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static string CrearInsumoSintetico(IReadOnlyList<object?[][]> grupos)
    {
        var filas = new List<object?[]>
        {
            new object?[] { "Pagos Pendientes por Aplicar", "Banco", "Fecha Recaudo", "Cantidad", "Valor" },
            new object?[] { null, null, null, null, "ASEO", "Total" },
            new object?[] { null, null, null, "66", "Total" }
        };

        decimal total = 0;
        foreach (var grupo in grupos)
        {
            foreach (var fila in grupo)
            {
                filas.Add(fila);
            }

            total += grupo[^1].OfType<decimal>().First();
        }

        filas.Add(["Resumen Recaudo Aplicado Por Servicio"]);
        filas.Add(["RECAUDO OCCIDENTE", "RECAUDO OTROS", "Total"]);
        filas.Add(["Total", null, total]);

        return CrearInsumo(filas);
    }

    /// <summary>
    /// Grupo de una fuente con la gramática medida: conceptos (etiqueta en la columna del concepto y la
    /// terna <c>D=E=F</c> a su derecha) y la fila de total rotulada con la etiqueta de la fuente.
    /// <paramref name="invertirConceptos"/> altera solo el ORDEN de las filas, nunca los montos.
    /// </summary>
    private static object?[][] Grupo(
        string fuente,
        decimal aplicadoR3,
        decimal anticiposR3 = 0m,
        decimal aplicadoR4 = 0m,
        decimal anticiposR4 = 0m,
        bool invertirConceptos = false)
    {
        var valores = new (string Etiqueta, decimal Valor)[]
        {
            ("1-APLICADOS A FACTURACION", aplicadoR3),
            ("2-SALDOS A FAVOR GENERADOS", anticiposR3),
            ("9-REVERSION PAGOS A COMPONENTES", aplicadoR4),
            ("9-REVERSION SALDOS A FAVOR", anticiposR4)
        };

        if (invertirConceptos)
        {
            Array.Reverse(valores);
        }

        var filas = new List<object?[]>();
        foreach (var (etiqueta, valor) in valores)
        {
            var fila = new object?[6];
            fila[2] = etiqueta;
            fila[3] = valor;
            fila[4] = valor;
            fila[5] = valor;
            filas.Add(fila);
        }

        var total = aplicadoR3 + anticiposR3 + aplicadoR4 + anticiposR4;
        var cierre = new object?[6];
        cierre[0] = fuente;
        cierre[1] = "Total";
        cierre[3] = total;
        cierre[4] = total;
        cierre[5] = total;
        filas.Add(cierre);

        return filas.ToArray();
    }

    /// <summary>Escribe un .xlsx mínimo (una hoja <c>Sheet1</c>) con las filas indicadas.</summary>
    private static string CrearInsumo(IReadOnlyList<object?[]> filas)
    {
        var ruta = Path.Combine(Path.GetTempPath(), $"p64-h5-insumo-{Guid.NewGuid():N}.xlsx");
        using var documento = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook);
        var parteWorkbook = documento.AddWorkbookPart();
        parteWorkbook.Workbook = new Workbook();
        var parteHoja = parteWorkbook.AddNewPart<WorksheetPart>();
        var datos = new SheetData();

        for (var i = 0; i < filas.Count; i++)
        {
            var fila = new Row { RowIndex = (uint)(i + 1) };
            var celdas = filas[i];
            for (var j = 0; j < celdas.Length; j++)
            {
                if (celdas[j] is null)
                {
                    continue;
                }

                var referencia = $"{LetraDeColumna(j + 1)}{i + 1}";
                if (celdas[j] is string texto)
                {
                    fila.Append(new Cell
                    {
                        CellReference = referencia,
                        DataType = CellValues.InlineString,
                        InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(texto))
                    });
                    continue;
                }

                fila.Append(new Cell
                {
                    CellReference = referencia,
                    CellValue = new CellValue(
                        Convert.ToDecimal(celdas[j], CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture))
                });
            }

            datos.Append(fila);
        }

        parteHoja.Worksheet = new Worksheet(datos);
        parteWorkbook.Workbook.AppendChild(new Sheets());
        parteWorkbook.Workbook.GetFirstChild<Sheets>()!.Append(new Sheet
        {
            Id = parteWorkbook.GetIdOfPart(parteHoja),
            SheetId = 1,
            Name = "Sheet1"
        });
        parteWorkbook.Workbook.Save();
        return ruta;
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Plantilla SINTÉTICA con la gramática de ANT EXT-REV en una geometría DESPLAZADA
    // (bloques en 40/60/80, columnas de valor D..M): evidencia de que el resolvedor LEYÓ la hoja.
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static readonly string[] ColumnasDeValor = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"];

    private static string CrearPlantillaDesplazada()
    {
        var ruta = Path.Combine(Path.GetTempPath(), $"p64-h5-geometria-{Guid.NewGuid():N}.xlsx");
        using var documento = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook);
        var parteWorkbook = documento.AddWorkbookPart();
        parteWorkbook.Workbook = new Workbook();
        var parteHoja = parteWorkbook.AddNewPart<WorksheetPart>();
        var datos = new SheetData();

        // Banda superior: dos filas de fórmula + dos controles booleanos (plan-then-write los exige).
        foreach (var numero in new uint[] { 2, 3 })
        {
            var fila = new Row { RowIndex = numero };
            foreach (var columna in ColumnasDeValor)
            {
                fila.Append(CeldaFormula($"{columna}{numero}", "0"));
            }

            datos.Append(fila);
        }

        var controles = new Row { RowIndex = 4 };
        controles.Append(CeldaBooleana("D4", "1=1"));
        controles.Append(CeldaBooleana("E4", "1=1"));
        datos.Append(controles);

        AgregarBloqueDesplazado(datos, 40, "BANCO DE OCCIDENTE");
        AgregarBloqueDesplazado(datos, 60, "EAAB RECIPROCIDAD");
        AgregarBloqueDesplazado(datos, 80, "ENEL");

        parteHoja.Worksheet = new Worksheet(datos);
        parteWorkbook.Workbook.AppendChild(new Sheets());
        parteWorkbook.Workbook.GetFirstChild<Sheets>()!.Append(new Sheet
        {
            Id = parteWorkbook.GetIdOfPart(parteHoja),
            SheetId = 1,
            Name = "ANT EXT-REV"
        });
        parteWorkbook.Workbook.Save();
        return ruta;
    }

    private static void AgregarBloqueDesplazado(SheetData datos, uint filaEtiqueta, string etiqueta)
    {
        var rotulo = new Row { RowIndex = filaEtiqueta };
        rotulo.Append(CeldaTexto($"B{filaEtiqueta}", etiqueta));
        datos.Append(rotulo);

        var ases = new Row { RowIndex = filaEtiqueta + 1 };
        var reportes = new Row { RowIndex = filaEtiqueta + 2 };
        for (var ase = 0; ase < 5; ase++)
        {
            var columna3 = ColumnasDeValor[ase * 2];
            var columna4 = ColumnasDeValor[(ase * 2) + 1];
            ases.Append(CeldaTexto($"{columna3}{filaEtiqueta + 1}", $"ASE {ase + 1}"));
            reportes.Append(CeldaTexto($"{columna3}{filaEtiqueta + 2}", "REPORTE 3"));
            reportes.Append(CeldaTexto($"{columna4}{filaEtiqueta + 2}", "REPORTE 4"));
        }

        ases.Append(CeldaTexto($"N{filaEtiqueta + 1}", "TOTAL"));
        datos.Append(ases);
        datos.Append(reportes);

        foreach (var (rotuloFila, numero) in new[] { ("APLICADO", filaEtiqueta + 3), ("ANTICIPOS", filaEtiqueta + 4) })
        {
            var fila = new Row { RowIndex = numero };
            fila.Append(CeldaTexto($"B{numero}", rotuloFila));
            foreach (var columna in ColumnasDeValor)
            {
                fila.Append(CeldaNumero($"{columna}{numero}", 0m));
            }

            datos.Append(fila);
        }

        var total = new Row { RowIndex = filaEtiqueta + 5 };
        total.Append(CeldaTexto($"B{filaEtiqueta + 5}", "TOTAL"));
        foreach (var columna in ColumnasDeValor)
        {
            total.Append(CeldaFormula($"{columna}{filaEtiqueta + 5}", $"{columna}{filaEtiqueta + 3}+{columna}{filaEtiqueta + 4}"));
        }

        datos.Append(total);
    }

    private static Cell CeldaTexto(string referencia, string texto) => new()
    {
        CellReference = referencia,
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(texto))
    };

    private static Cell CeldaNumero(string referencia, decimal valor) => new()
    {
        CellReference = referencia,
        CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture))
    };

    private static Cell CeldaFormula(string referencia, string texto) => new()
    {
        CellReference = referencia,
        CellFormula = new CellFormula(texto),
        CellValue = new CellValue("0"),
        DataType = CellValues.Number
    };

    private static Cell CeldaBooleana(string referencia, string texto) => new()
    {
        CellReference = referencia,
        CellFormula = new CellFormula(texto),
        CellValue = new CellValue("1"),
        DataType = CellValues.Boolean
    };

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
}
