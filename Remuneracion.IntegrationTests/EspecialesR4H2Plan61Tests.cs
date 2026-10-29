using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// PLAN 61 — H2/F2: regresiones de <b>valores medidos</b> y de <b>forma</b> de los Especiales del R4.
///
/// <para><b>Doctrina del fixture (H5.4).</b> Ningún aserto es la identidad de su propia fórmula
/// (<c>D9 - P9 == TotalReversionEsperada</c> está prohibido). Cada valor esperado viaja con su
/// derivación medida: fuente R4 (por <b>encabezado</b>), administrativo y R10 (por <b>etiqueta</b>
/// — el ASE se localiza por su identificador/rótulo, nunca por número de fila). Los tres oráculos
/// se leen con BCL puro sobre el libro; la fuente R4 se lee con <see cref="ExcelDataReader"/> sobre
/// una copia temporal (Excel cerrado).</para>
///
/// <para><b>Cobertura.</b></para>
/// <list type="bullet">
/// <item>H2.1: 2026091 ASE 2 (Δ P = −26.500,00; DetRetri 14.147.357.755), 2026092 ASE 4
/// (Δ P = −49.900,44; DetRetri 7.675.362.502) y 2026091 ASE 3 — el «caso del peso» (§13.2-C):
/// D104:D108 = 10.891.338.739,48 → ROUND 739 vs R10 740 (Δ 1). El R10 es errata humana y el test
/// lo documenta para que nadie «arregle» el redondeo mañana.</item>
/// <item>H2.2: «sin columna → 0» por AUSENCIA DECLARADA (no por literal) y «columna movida de
/// letra» (misma fuente, encabezado en T en vez de P ⇒ mismo resultado).</item>
/// </list>
/// </summary>
public sealed class EspecialesR4H2Plan61Tests
{
    static EspecialesR4H2Plan61Tests()
    {
        // ExcelReaderConfiguration crea un Encoding de la página 1252: sin el proveedor el reader
        // falla según el orden en que la suite haya tocado ExcelDataReader.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <summary>Tolerancia de comparación: el centavo (las fuentes traen dos decimales).</summary>
    private const decimal Centavo = 0.005m;

    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    private const string HojaR4 = "Reversion Pagos R4";

    private static readonly string[] AliasesEspeciales = ["SERVICIO ESPECIALES", "Especiales"];

    private static string Raiz => RaizRepositorio();

    private static string RutaPlantillaV4 => Path.Combine(
        Raiz, "Docs", "Plantilla estandarizada", "v4", "Formato Remuneracion 2026xx-x Total v4.xlsx");

    private static string CarpetaSeptiembre1 => Path.Combine(Raiz, "Docs", "Prueba Septiembre 1", "Insumos");
    private static string CarpetaSeptiembre2 => Path.Combine(Raiz, "Docs", "Prueba Septiembre-2", "Insumos");

    private static string RutaAdminSeptiembre1 => Path.Combine(
        Raiz, "Docs", "Prueba Septiembre 1", "Resultados", "Administrativo", "Remuneracion 202609-1 Total.xlsx");

    private static string RutaAdminSeptiembre2 => Path.Combine(
        Raiz, "Docs", "Prueba Septiembre-2", "Resultados", "Administrativo", "Remuneracion 202609-2 Total.xlsx");

    private static string RutaR10Septiembre1 => Path.Combine(CarpetaSeptiembre1, "R10_Remuneracion_2026091.xlsx");
    private static string RutaR10Septiembre2 => Path.Combine(CarpetaSeptiembre2, "R10_Remuneracion_2026092.xlsx");

    private static Periodo PeriodoSeptiembre1 => new() { CodigoAAAAMM = "202609", NumeroQuincena = 1 };
    private static Periodo PeriodoSeptiembre2 => new() { CodigoAAAAMM = "202609", NumeroQuincena = 2 };

    /// <summary>
    /// Corrida de CINCO ASE de 2026092, compartida. La quincena 2 exige el camino multi-ASE: el camino
    /// de un solo ASE no lee SALDOS POR NOTA / RETRIBUCIÓN NEGATIVA y por tanto no puede reproducir el
    /// total del ASE. La corrida está medida en §14.3 (exit 0, <c>CONSOLIDADO!D107 = 7.675.362.501,96</c>).
    /// </summary>
    internal static readonly Lazy<ResultadoProcesoPeriodo> Generacion2026092 = new(() =>
    {
        var salida = Path.Combine(CarpetaTemporal("p61-h2-2026092"), "Remuneracion 202609-2.xlsx");
        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());

        return procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = PeriodoSeptiembre2,
            CarpetaPeriodo = CarpetaSeptiembre2,
            RutaPlantilla = RutaPlantillaV4,
            RutaSalida = salida
        });
    });

    /// <summary>
    /// Corrida de CINCO ASE de 2026091, compartida por los tests de ASE 2 y ASE 3.
    ///
    /// <para><b>Por qué ya no hay validador decorador.</b> Hasta H2 este fixture omitía el gate
    /// DetRetri-vs-R10 con un decorador, porque ese gate era fail-fast ANTES de escribir y el período
    /// no cierra por la errata humana del ASE 3 (Δ 1 peso, §13.2-C). Plan 61 (H3) lo convirtió en
    /// NOVEDAD post-escritura: el procesador devuelve el resultado con el Excel marcado y los leafs
    /// quedan disponibles sin relajar ninguna aserción. El Δ sigue documentado en
    /// <c>H21_2026091_Ase3_ElCasoDelPeso_R10EsErrataHumana</c>.</para>
    /// </summary>
    internal static readonly Lazy<(ResultadoProcesoPeriodo Resultado, IReadOnlyList<WorkbookLeafInputs> Leafs)> Generacion2026091 = new(() =>
    {
        var salida = Path.Combine(CarpetaTemporal("p61-h2-2026091"), "Remuneracion 202609-1.xlsx");
        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader(),
            new ValidacionOracleReader(),
            new ReporteDivergenciasDeclaradasWriter(),
            new ReporteNovedadesWriter());

        var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = PeriodoSeptiembre1,
            CarpetaPeriodo = CarpetaSeptiembre1,
            RutaPlantilla = RutaPlantillaV4,
            RutaSalida = salida
        });

        return (resultado, resultado.Leafs);
    });

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // H2.1 — regresiones con valores MEDIDOS (por encabezado / por etiqueta)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 2026091 ASE 2 (Lime, primera quincena). Fuente R4: cierre D = −8.739.602,37, P
    /// («SERVICIO ESPECIALES» por encabezado) = −26.500,00 ⇒ visible de la plantilla
    /// <c>D97−P97 = −8.713.102,37</c>. El total D104:D108 del ASE es 14.147.357.755,39 y el
    /// DetRetri (ROUND) 14.147.357.755, que es exactamente el R10 <c>DetRetri2026091</c> del ASE 2.
    ///
    /// <para><b>Actualización Plan 62 (H5, T5.2 — data drift R10):</b> el R10 vigente del período
    /// expresa los montos como TEXTO MONETARIO en <c>sharedStrings</c> (<c>t="s"</c>, p. ej.
    /// <c>"$14.147.357.755"</c>) y no como resultado de fórmula (<c>t="str"</c>). La narrativa que
    /// este fixture blinda no es la representación de la celda, sino que el R10 es una COPIA
    /// FORMATEADA de la composición (no un cálculo independiente): por eso se exige que su valor
    /// coincida al peso con el <c>Detalle</c> del dominio y con el CONSOLIDADO del administrativo,
    /// sin fijar el vehículo (texto de fórmula, texto compartido o número).</para>
    /// </summary>
    [Fact]
    public void H21_2026091_Ase2_R4EspecialesYDetRetri_ValoresMedidos()
    {
        var (_, _, r4) = RutasAse(Path.Combine(CarpetaSeptiembre1, "2-Lime"));

        var fuente = MedirR4(r4);
        Assert.Equal("SERVICIO ESPECIALES", fuente.EncabezadoEspeciales);
        AfirmarMonto("2026091 ASE2 · R4 cierre · columna Total (D)", -8739602.37m, fuente.D);
        AfirmarMonto("2026091 ASE2 · R4 cierre · SERVICIO ESPECIALES (P, por encabezado)", -26500.00m, fuente.Especiales);

        var agregado = new ExcelDataReaderRecaudoReader().LeerR4(r4);
        AfirmarMonto("2026091 ASE2 · agregado LeerR4 · TotalReversiones (D observado)", -8739602.37m, agregado.TotalReversiones);
        AfirmarMonto("2026091 ASE2 · agregado LeerR4 · EspecialesReversion (por encabezado)", -26500.00m, agregado.EspecialesReversion);

        // Corrida multi-ASE de 2026091 (el validador decorador captura los leafs; el gate R10 se
        // ejerce aparte con los oráculos medidos). El DetRetri sale de la composición del dominio
        // que la plantilla publica en D104:D108 — no de un agregado que pueda divergir de la hoja.
        var (_, leafsSeptiembre1) = Generacion2026091.Value;
        var leaf = leafsSeptiembre1.Single(l => l.Ase.Id == 2);
        AfirmarMonto("2026091 ASE2 · leaf R4 · D9", -8739602.37m, leaf.R4.D9);
        AfirmarMonto("2026091 ASE2 · leaf R4 · P9", -26500.00m, leaf.R4.P9);
        AfirmarMonto("2026091 ASE2 · leaf R4 · TotalReversionEsperada", -8713102.37m, leaf.R4.TotalReversionEsperada);

        var detRetri = leaf.DetRetriQ2 ?? throw new InvalidOperationException("ASE 2: falta DetRetriQ2 en el leaf.");
        AfirmarMonto("2026091 ASE2 · CONSOLIDADO D104:D108 (TotalD104)", 14147357755.39m, detRetri.TotalD104);
        Assert.Equal(14147357755m, detRetri.Detalle);

        AfirmarMonto("2026091 ASE2 · administrativo CONSOLIDADO · recaudo total del ASE", 14147357755.39m, ConsolidadoDPorAse(RutaAdminSeptiembre1, 2));

        var r10 = R10DetRetriPorAse(RutaR10Septiembre1, "2026091", 2);
        Assert.True(
            r10.EsTexto,
            "El R10 del período es una copia formateada de la composición (texto tecleado o texto "
            + "monetario en sharedStrings), no un cálculo: su valor no es un oráculo independiente.");
        AfirmarMonto("2026091 ASE2 · R10 DetRetri2026091 · recaudo total del ASE", 14147357755m, r10.Valor);
        Assert.Equal(r10.Valor, detRetri.Detalle);
    }

    /// <summary>
    /// 2026092 ASE 4 (Bogotá Limpia, segunda quincena). Fuente R4: cierre D = −178.619.709,53, P
    /// («SERVICIO ESPECIALES») = −49.900,44 ⇒ visible <c>D239−P239 = −178.569.809,09</c>. El total
    /// D104:D108 es 7.675.362.501,96 y el DetRetri 7.675.362.502 = R10 <c>DetRetri2026092</c> ASE 4.
    /// </summary>
    [Fact]
    public void H21_2026092_Ase4_R4EspecialesYDetRetri_ValoresMedidos()
    {
        var ase = AseFactory.DesdeId(4);
        var (r1, r2, r4) = RutasAse(Path.Combine(CarpetaSeptiembre2, "4-Bogota Limpia"));

        var fuente = MedirR4(r4);
        Assert.Equal("SERVICIO ESPECIALES", fuente.EncabezadoEspeciales);
        AfirmarMonto("2026092 ASE4 · R4 cierre · columna Total (D)", -178619709.53m, fuente.D);
        AfirmarMonto("2026092 ASE4 · R4 cierre · SERVICIO ESPECIALES (P, por encabezado)", -49900.44m, fuente.Especiales);

        var agregado = new ExcelDataReaderRecaudoReader().LeerR4(r4);
        AfirmarMonto("2026092 ASE4 · agregado LeerR4 · TotalReversiones (D observado)", -178619709.53m, agregado.TotalReversiones);
        AfirmarMonto("2026092 ASE4 · agregado LeerR4 · EspecialesReversion (por encabezado)", -49900.44m, agregado.EspecialesReversion);

        // La corrida multi-ASE es la única que puede reproducir la segunda quincena (ajustes SF-T).
        var resultado = Generacion2026092.Value;
        var leaf = resultado.Leafs.Single(l => l.Ase.Id == 4);
        AfirmarMonto("2026092 ASE4 · leaf R4 · D9", -178619709.53m, leaf.R4.D9);
        AfirmarMonto("2026092 ASE4 · leaf R4 · P9", -49900.44m, leaf.R4.P9);
        AfirmarMonto("2026092 ASE4 · leaf R4 · TotalReversionEsperada", -178569809.09m, leaf.R4.TotalReversionEsperada);

        AfirmarMonto(
            "2026092 ASE4 · administrativo 'Reversion Pagos R4' · visible del bloque",
            -178569809.09m,
            AdminVisibleR4(RutaAdminSeptiembre2, ase.NombreCorto));

        var detRetri = leaf.DetRetriQ2 ?? throw new InvalidOperationException("ASE 4: falta DetRetriQ2 en el leaf.");
        AfirmarMonto("2026092 ASE4 · CONSOLIDADO D104:D108 (TotalD104)", 7675362501.96m, detRetri.TotalD104);
        Assert.Equal(7675362502m, detRetri.Detalle);
        Assert.True(File.Exists(resultado.RutaSalida), $"La corrida de 2026092 debía certificar la salida: {resultado.RutaSalida}");

        AfirmarMonto("2026092 ASE4 · administrativo CONSOLIDADO · recaudo total del ASE", 7675362501.96m, ConsolidadoDPorAse(RutaAdminSeptiembre2, 4));

        var r10 = R10DetRetriPorAse(RutaR10Septiembre2, "2026092", 4);
        AfirmarMonto("2026092 ASE4 · R10 DetRetri2026092 · recaudo total del ASE", 7675362502m, r10.Valor);
        Assert.Equal(r10.Valor, detRetri.Detalle);
    }

    /// <summary>
    /// 2026091 ASE 3 — el «caso del peso» (§13.2-C). El administrativo calcula
    /// <c>D106 = 10.891.338.739,48</c> como suma exacta de sus cinco bloques y la regla única de
    /// redondeo (<see cref="DetRetriRounder"/>, ROUND de Excel) da <b>10.891.338.739</b> (últimos
    /// dígitos <b>739</b>).
    ///
    /// <para><b>Estado vigente del dato (data drift, actualizado en Plan 62 H5/T5.2):</b> el R10
    /// REEMITIDO del período ya trae el valor correcto <b>739</b> (la errata humana 740 quedó
    /// corregida en el dato); el que conserva el <b>740</b> es el DetRetri del manual del
    /// administrativo. Este test fija AMBOS números para que nadie «corrija» el redondeo ni amplíe
    /// la tolerancia: el Δ = 1 peso que queda entre la app (739) y el manual (740) es ERRATA DEL
    /// MANUAL, declarada en el comparador como <c>P62-DETRETRI-D11-ERRATA-PESO-ASE3</c> (y su
    /// arrastre al TOTAL en <c>P62-DETRETRI-D14-TOTAL-ERRATA-PESO</c>, H5/T5.1).</para>
    /// </summary>
    [Fact]
    public void H21_2026091_Ase3_ElCasoDelPeso_ErrataHumanaCorregidaEnElR10()
    {
        var (_, leafsSeptiembre1) = Generacion2026091.Value;
        var leaf = leafsSeptiembre1.Single(l => l.Ase.Id == 3);

        var detRetri = leaf.DetRetriQ2 ?? throw new InvalidOperationException("ASE 3: falta DetRetriQ2 en el leaf.");
        AfirmarMonto("2026091 ASE3 · CONSOLIDADO D104:D108 (TotalD104)", 10891338739.48m, detRetri.TotalD104);
        Assert.Equal(10891338739m, detRetri.Detalle);

        AfirmarMonto(
            "2026091 ASE3 · administrativo CONSOLIDADO · recaudo total del ASE",
            10891338739.48m,
            ConsolidadoDPorAse(RutaAdminSeptiembre1, 3));

        // El DetRetri del administrativo conserva el 740 humano; el R10 vigente ya trae el 739 del
        // ROUND. Ninguno de los dos es un cálculo independiente (los dos copian la composición).
        var admin = R10DetRetriPorAse(RutaAdminSeptiembre1, "2026091", 3);
        AfirmarMonto("2026091 ASE3 · administrativo DetRetri2026091 · ASE 3 (errata: 740)", 10891338740m, admin.Valor);

        var r10 = R10DetRetriPorAse(RutaR10Septiembre1, "2026091", 3);
        Assert.True(
            r10.EsTexto,
            "El R10 del período es una copia formateada de la composición (texto monetario), no un oráculo independiente.");
        AfirmarMonto("2026091 ASE3 · R10 DetRetri2026091 · ASE 3 (corregido: 739)", 10891338739m, r10.Valor);
        Assert.Equal(r10.Valor, detRetri.Detalle);

        // Δ = 1 peso exacto, y vive en el MANUAL: el redondeo de la plantilla (739) NO se toca y la
        // tolerancia NO se amplía.
        Assert.Equal(1m, admin.Valor - detRetri.Detalle);
        Assert.Equal(0m, r10.Valor - detRetri.Detalle);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // H2.2 — regresiones de FORMA
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// «Sin columna → 0» por <b>ausencia declarada</b>, no por literal. Se miden dos formas de
    /// ausencia:
    /// <list type="bullet">
    /// <item>2026091 ASE 1: la fuente simplemente no trae el encabezado.</item>
    /// <item>2026092 ASE 2: en la letra P trae <c>Ajuste Recargo Aseo</c> con −10,16 en el cierre
    /// (medido). Un lector posicional devolvería −10,16; el producto debe devolver 0 porque el
    /// <b>encabezado</b> de Especiales no existe.</item>
    /// </list>
    /// El 0 es la ausencia medida, nunca un valor inventado.
    /// </summary>
    [Fact]
    public void H22_SinEncabezadoDeEspeciales_ElValorEsCeroPorAusenciaDeclarada()
    {
        // Caso 1: ausencia simple (2026091 ASE 1).
        var ase1 = AseFactory.DesdeId(1);
        var (r1Ase1, r2Ase1, r4Ase1) = RutasAse(Path.Combine(CarpetaSeptiembre1, "1-Promoambiental"));
        var fuente1 = MedirR4(r4Ase1);
        Assert.Null(fuente1.EncabezadoEspeciales);
        Assert.Null(fuente1.LetraEspeciales);

        var agregado1 = new ExcelDataReaderRecaudoReader().LeerR4(r4Ase1);
        AfirmarMonto("2026091 ASE1 · sin encabezado · TotalReversiones = D observado", -3544507.29m, agregado1.TotalReversiones);
        Assert.Equal(0m, agregado1.EspecialesReversion);

        var leaf1 = LeerLeaf(ase1, PeriodoSeptiembre1, r1Ase1, r2Ase1, r4Ase1);
        Assert.Equal(0m, leaf1.R4.P9);
        AfirmarMonto("2026091 ASE1 · sin encabezado · TotalReversionEsperada = D (neto sin especiales)", -3544507.29m, leaf1.R4.TotalReversionEsperada);

        // Caso 2: la letra P EXISTE pero porta otro encabezado (2026092 ASE 2). La ausencia es del
        // ENCABEZADO, no de la columna física: el lector no puede depender de la posición.
        var ase2 = AseFactory.DesdeId(2);
        var (r1Ase2, r2Ase2, r4Ase2) = RutasAse(Path.Combine(CarpetaSeptiembre2, "2-Lime"));
        var fuente2 = MedirR4(r4Ase2);
        Assert.Null(fuente2.EncabezadoEspeciales);
        Assert.Null(fuente2.LetraEspeciales);
        Assert.Equal("Ajuste Recargo Aseo", fuente2.Encabezados.Single(e => e.Letra == "P").Encabezado);
        AfirmarMonto("2026092 ASE2 · P del cierre (Ajuste Recargo Aseo) — lo que un lector posicional inyectaría", -10.16m, fuente2.ValoresCierre["P"] ?? decimal.Zero);

        var agregado2 = new ExcelDataReaderRecaudoReader().LeerR4(r4Ase2);
        AfirmarMonto("2026092 ASE2 · sin encabezado · TotalReversiones = D observado", -45230052m, agregado2.TotalReversiones);
        Assert.Equal(0m, agregado2.EspecialesReversion);

        var leaf2 = LeerLeaf(ase2, PeriodoSeptiembre2, r1Ase2, r2Ase2, r4Ase2);
        Assert.Equal(0m, leaf2.R4.P9);
        AfirmarMonto("2026092 ASE2 · sin encabezado · TotalReversionEsperada = D (no −10,16)", -45230052m, leaf2.R4.TotalReversionEsperada);
    }

    /// <summary>
    /// «Columna movida de letra»: se toma una COPIA de la misma fuente de 2026092 ASE 4 y se mueve
    /// físicamente la columna <c>SERVICIO ESPECIALES</c> de P a T (encabezado y todos sus valores,
    /// incluido el cierre). El resultado debe ser idéntico: si el lector dependiera de la letra P,
    /// este test caería. Es la regresión anti-M-B del §2-D-H1.
    /// </summary>
    [Fact]
    public void H22_EncabezadoEnOtraLetra_MismoResultado_SinDependenciaPosicional()
    {
        var ase = AseFactory.DesdeId(4);
        var (r1, r2, r4Original) = RutasAse(Path.Combine(CarpetaSeptiembre2, "4-Bogota Limpia"));

        var fuenteOriginal = MedirR4(r4Original);
        Assert.Equal("P", fuenteOriginal.LetraEspeciales);

        var leafOriginal = LeerLeaf(ase, PeriodoSeptiembre2, r1, r2, r4Original);
        var agregadoOriginal = new ExcelDataReaderRecaudoReader().LeerR4(r4Original);

        var carpetaCopia = CarpetaTemporal("p61-h2-columna-movida");
        var r4Movido = Path.Combine(carpetaCopia, Path.GetFileName(r4Original));
        File.Copy(r4Original, r4Movido);
        MoverColumna(r4Movido, desde: "P", hacia: "T");

        var fuenteMovida = MedirR4(r4Movido);
        Assert.Equal("SERVICIO ESPECIALES", fuenteMovida.EncabezadoEspeciales);
        Assert.Equal("T", fuenteMovida.LetraEspeciales);
        AfirmarMonto("2026092 ASE4 (columna movida P→T) · cierre · Especiales", -49900.44m, fuenteMovida.Especiales);

        var leafMovido = LeerLeaf(ase, PeriodoSeptiembre2, r1, r2, r4Movido);
        AfirmarMonto("columna movida · D9 idéntico", leafOriginal.R4.D9, leafMovido.R4.D9);
        AfirmarMonto("columna movida · P9 idéntico", leafOriginal.R4.P9, leafMovido.R4.P9);
        AfirmarMonto("columna movida · TotalReversionEsperada idéntico", leafOriginal.R4.TotalReversionEsperada, leafMovido.R4.TotalReversionEsperada);

        var agregadoMovido = new ExcelDataReaderRecaudoReader().LeerR4(r4Movido);
        AfirmarMonto("columna movida · agregado TotalReversiones idéntico", agregadoOriginal.TotalReversiones, agregadoMovido.TotalReversiones);
        AfirmarMonto("columna movida · agregado EspecialesReversion idéntico", agregadoOriginal.EspecialesReversion, agregadoMovido.EspecialesReversion);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Cadena fuente → lectores → leaf (sin composición propia): el total que publica D104:D108
    // sale de la corrida real, nunca de una fórmula re-escrita en el test.
    // ════════════════════════════════════════════════════════════════════════════════════════════

    private static WorkbookLeafInputs LeerLeaf(Ase ase, Periodo periodo, string r1, string r2, string r4) =>
        new ExcelDataReaderWorkbookLeafInputReader { RutaPlantilla = RutaPlantillaV4 }
            .LeerLeafInputs(ase, periodo, r1, r2, r4);

    private static (string R1, string R2, string R4) RutasAse(string carpetaAse)
    {
        var locator = new ArchivoFuenteLocator();
        var r1 = locator.BuscarArchivo(carpetaAse, "Recaudoporcomponente")
            ?? throw new FileNotFoundException($"Falta R1 en '{carpetaAse}'.");
        var r2 = locator.BuscarArchivo(carpetaAse, "RerpoteDetalleSaldosaFavor")
            ?? throw new FileNotFoundException($"Falta R2 en '{carpetaAse}'.");
        var r4 = locator.BuscarArchivo(carpetaAse, "ReversiónPorComponente")
            ?? locator.BuscarArchivo(carpetaAse, "ReversionPorComponente")
            ?? throw new FileNotFoundException($"Falta R4 en '{carpetaAse}'.");

        return (r1, r2, r4);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Lectura de oráculos POR ETIQUETA (BCL puro: ZipArchive + XLinq)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Recaudo total de un ASE en el CONSOLIDADO del administrativo. El ASE se localiza por su
    /// identificador en la columna <c>B</c> del bloque de datos (encabezado «ASE»), nunca por fila.
    /// </summary>
    private static decimal ConsolidadoDPorAse(string rutaLibro, int aseId)
    {
        var libro = LeerLibro(rutaLibro);
        var filas = libro.Filas[HojaConsolidado].ToList();
        // La hoja repite el bloque «ASE/Concesionarios» (un resumen arriba y el consolidado final
        // abajo). El valor que publica el período es el del ÚLTIMO bloque — se localiza por rótulo,
        // nunca por número de fila.
        var inicios = Enumerable.Range(0, filas.Count)
            .Where(i => IgualEtiqueta(filas[i].Texto("B"), "ASE")
                && filas[i].Texto("C").StartsWith("Concesionarios", StringComparison.OrdinalIgnoreCase)
                && filas[i].Texto("F").StartsWith("Recaudo por tratamiento de lixiviados", StringComparison.OrdinalIgnoreCase))
            .ToList();
        Assert.NotEmpty(inicios);
        var inicio = inicios[^1];

        var fila = filas.Skip(inicio + 1).FirstOrDefault(f => f.Texto("B") == aseId.ToString(CultureInfo.InvariantCulture))
            ?? throw new InvalidOperationException($"'{Path.GetFileName(rutaLibro)}': CONSOLIDADO sin fila para el ASE {aseId}.");
        return Monto(fila.Celda("D") ?? throw new InvalidOperationException($"ASE {aseId}: CONSOLIDADO sin celda D."));
    }

    /// <summary>
    /// Visible del bloque R4 del ASE en el administrativo: la fila «TOTAL OPORTUNO» que sigue al
    /// rótulo del ASE (columna B = nombre corto, p. ej. <c>BOGOTA LIMPIA</c>). Su D es la fórmula
    /// <c>D{total}−P{total}</c>: el valor público del bloque, no el de la app.
    /// </summary>
    private static decimal AdminVisibleR4(string rutaLibro, string nombreCorto)
    {
        var libro = LeerLibro(rutaLibro);
        var filas = libro.Filas[HojaR4].ToList();
        var inicio = filas.FindIndex(f => IgualEtiqueta(f.Texto("B"), nombreCorto));
        Assert.True(inicio >= 0, $"'{Path.GetFileName(rutaLibro)}': 'Reversion Pagos R4' sin rótulo '{nombreCorto}'.");

        for (var i = inicio; i < filas.Count; i++)
        {
            var fila = filas[i];
            if (IgualEtiqueta(fila.Texto("A"), "TOTAL") && IgualEtiqueta(fila.Texto("B"), "OPORTUNO"))
            {
                return Monto(fila.Celda("D") ?? throw new InvalidOperationException($"'{nombreCorto}': TOTAL OPORTUNO sin D."));
            }
        }

        throw new InvalidOperationException($"'{Path.GetFileName(rutaLibro)}': bloque '{nombreCorto}' sin fila TOTAL OPORTUNO.");
    }

    /// <summary>
    /// Fila del DetRetri del período/R10 correspondiente al ASE, localizada por el identificador de
    /// la columna <c>B</c>. Devuelve el valor de <c>D</c> y si la celda es TEXTO (tecleado
    /// <c>t="str"</c> o texto monetario en <c>sharedStrings</c> <c>t="s"</c>), condición que
    /// descalifica al R10 como oráculo independiente (§13.2). Plan 62 (H5, T5.2 — data drift): el R10
    /// vigente del período usa <c>t="s"</c>; ambas formas son texto formateado, no un cálculo.
    /// </summary>
    private static (decimal Valor, bool EsTexto) R10DetRetriPorAse(string rutaLibro, string periodo, int aseId)
    {
        var libro = LeerLibro(rutaLibro);
        var hoja = $"DetRetri{periodo}";
        Assert.True(libro.Filas.TryGetValue(hoja, out var filasHoja), $"'{Path.GetFileName(rutaLibro)}' no trae la hoja '{hoja}'.");

        var filas = filasHoja!.ToList();
        // La hoja puede repetir el rótulo ASE en varios bloques (DetRetri, diferencias, validación);
        // el bloque del DetRetri es el que declara D = «RECAUDO TOTAL».
        var inicio = filas.FindIndex(f =>
            IgualEtiqueta(f.Texto("B"), "ASE") && IgualEtiqueta(f.Texto("D"), "RECAUDO TOTAL"));
        Assert.True(inicio >= 0, $"'{hoja}': no se encontró el bloque ASE/RECAUDO TOTAL.");

        var fila = filas.Skip(inicio + 1)
            .FirstOrDefault(f => f.Texto("B") == aseId.ToString(CultureInfo.InvariantCulture) && f.Celda("D") is not null)
            ?? throw new InvalidOperationException($"'{hoja}': sin fila DetRetri para el ASE {aseId}.");

        var celda = fila.Celda("D")!;
        return (Monto(celda), EsCeldaDeTexto(celda.Tipo));
    }

    /// <summary>La celda guarda TEXTO (tecleado, compartido o inline), no un número.</summary>
    private static bool EsCeldaDeTexto(string? tipo) => tipo is "s" or "str" or "inlineStr";

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Medición de la fuente R4 (ExcelDataReader, solo-lectura, sobre copia temporal)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    private sealed record EncabezadoFisico(string Letra, string Encabezado);

    private sealed record MedidaR4(
        int FilaEncabezados,
        IReadOnlyList<EncabezadoFisico> Encabezados,
        string? LetraEspeciales,
        string? EncabezadoEspeciales,
        decimal D,
        decimal Especiales,
        IReadOnlyDictionary<string, decimal?> ValoresCierre);

    /// <summary>
    /// Mide la fuente R4: fila de encabezados, encabezado de Especiales (por alias normalizado, nunca
    /// por índice), fila de cierre (<c>A="Total"</c> con <c>B</c> vacío) y sus valores por letra.
    /// </summary>
    private static MedidaR4 MedirR4(string ruta)
    {
        Assert.True(File.Exists(ruta), $"Falta la fuente R4: {ruta}");

        var carpeta = CarpetaTemporal("p61-h2-fuente");
        var copia = Path.Combine(carpeta, Path.GetFileName(ruta));
        File.Copy(ruta, copia);
        try
        {
            using var stream = new FileStream(copia, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var lector = ExcelReaderFactory.CreateReader(stream);

            var filas = new List<object?[]>();
            while (lector.Read())
            {
                var fila = new object?[lector.FieldCount];
                for (var i = 0; i < fila.Length; i++)
                {
                    fila[i] = lector.GetValue(i);
                }

                filas.Add(fila);
            }

            var indiceEncabezados = filas.FindIndex(f =>
                f.Count(v => !string.IsNullOrWhiteSpace(Texto(v))) >= 5
                && f.Any(v => Texto(v).Equals("Total", StringComparison.OrdinalIgnoreCase)));
            Assert.True(indiceEncabezados >= 0, $"'{Path.GetFileName(ruta)}' no trae fila de encabezados reconocible.");

            var encabezados = filas[indiceEncabezados]
                .Select((v, i) => (Letra: Letra(i), Texto: Texto(v)))
                .Where(e => e.Texto.Length > 0)
                .Select(e => new EncabezadoFisico(e.Letra, e.Texto))
                .ToList();

            var indicesCierre = filas
                .Select((f, i) => (Fila: f, Indice: i))
                .Where(x => Texto(x.Fila.ElementAtOrDefault(0)).Equals("Total", StringComparison.OrdinalIgnoreCase)
                    && string.IsNullOrWhiteSpace(Texto(x.Fila.ElementAtOrDefault(1))))
                .Select(x => x.Indice)
                .ToList();
            Assert.Single(indicesCierre);

            var filaCierre = filas[indicesCierre[0]];
            var valoresCierre = filaCierre
                .Select((v, i) => (Letra: Letra(i), Valor: NumeroONulo(v)))
                .ToDictionary(x => x.Letra, x => x.Valor, StringComparer.Ordinal);

            var especiales = encabezados.FirstOrDefault(e => EsEspeciales(e.Encabezado));
            var d = NumeroONulo(filaCierre.ElementAtOrDefault(3))
                ?? throw new InvalidOperationException($"'{Path.GetFileName(ruta)}': el cierre no trae monto en Total.");

            return new MedidaR4(
                indiceEncabezados + 1,
                encabezados,
                especiales?.Letra,
                especiales?.Encabezado,
                d,
                especiales is null ? 0m : NumeroONulo(filaCierre.ElementAtOrDefault(IndiceDeLetra(especiales.Letra))) ?? 0m,
                valoresCierre);
        }
        finally
        {
            Directory.Delete(carpeta, recursive: true);
        }
    }

    /// <summary>Mueve físicamente una columna completa (encabezado + valores) dentro de la hoja.</summary>
    private static void MoverColumna(string rutaCopia, string desde, string hacia)
    {
        using var documento = SpreadsheetDocument.Open(rutaCopia, true);
        var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null.");
        var libro = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null.");
        var hoja = libro.Descendants<Sheet>().Single();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(hoja.Id!);
        var worksheet = worksheetPart.Worksheet
            ?? throw new InvalidOperationException("La hoja no trae Worksheet.");
        var sheetData = worksheet.GetFirstChild<SheetData>()
            ?? throw new InvalidOperationException("La hoja no trae SheetData.");

        var filaMaxima = 0;
        foreach (var row in sheetData.Elements<Row>())
        {
            filaMaxima = Math.Max(filaMaxima, (int)(row.RowIndex?.Value ?? 0u));
            var celda = row.Elements<Cell>()
                .FirstOrDefault(c => string.Equals(LetraDeReferencia(c.CellReference?.Value), desde, StringComparison.Ordinal));
            if (celda is null)
            {
                continue;
            }

            celda.CellReference = hacia + row.RowIndex!.Value.ToString(CultureInfo.InvariantCulture);

            // Excel exige celdas en orden ascendente de columna dentro de la fila.
            var ordenadas = row.Elements<Cell>()
                .OrderBy(c => IndiceDeLetra(LetraDeReferencia(c.CellReference?.Value)))
                .ToList();
            row.RemoveAllChildren<Cell>();
            foreach (var ordenada in ordenadas)
            {
                row.Append(ordenada);
            }
        }

        var dimension = worksheet.GetFirstChild<SheetDimension>();
        if (dimension is not null)
        {
            dimension.Reference = $"A1:{hacia}{filaMaxima.ToString(CultureInfo.InvariantCulture)}";
        }

        worksheet.Save();
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Lectura BCL de libros (administrativo / R10)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    private sealed record CeldaXlsx(string Referencia, string? Tipo, string? Formula, string? Valor, string? Texto);

    private sealed record FilaXlsx(int Numero, IReadOnlyDictionary<string, CeldaXlsx> Celdas)
    {
        public CeldaXlsx? Celda(string letra) => Celdas.GetValueOrDefault(letra);

        public string Texto(string letra)
        {
            var celda = Celda(letra);
            return (celda?.Texto ?? celda?.Valor ?? string.Empty).Trim();
        }
    }

    private sealed record LibroXlsx(IReadOnlyDictionary<string, IReadOnlyList<FilaXlsx>> Filas);

    private static LibroXlsx LeerLibro(string ruta)
    {
        Assert.True(File.Exists(ruta), $"Falta el libro: {ruta}");

        using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);

        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace relacionesNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace paqueteNs = "http://schemas.openxmlformats.org/package/2006/relationships";

        var compartidas = new List<string>();
        var entradaCadenas = zip.GetEntry("xl/sharedStrings.xml");
        if (entradaCadenas is not null)
        {
            using var cadenas = entradaCadenas.Open();
            foreach (var si in XDocument.Load(cadenas).Root!.Elements(ns + "si"))
            {
                compartidas.Add(string.Concat(si.Descendants(ns + "t").Select(t => t.Value)));
            }
        }

        XDocument libro;
        using (var s = (zip.GetEntry("xl/workbook.xml") ?? throw new InvalidOperationException($"'{ruta}' no es un xlsx legible.")).Open())
        {
            libro = XDocument.Load(s);
        }

        XDocument relaciones;
        using (var s = zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open())
        {
            relaciones = XDocument.Load(s);
        }

        var destinos = relaciones.Root!.Elements(paqueteNs + "Relationship")
            .ToDictionary(r => r.Attribute("Id")!.Value, r => r.Attribute("Target")!.Value);

        var hojas = new Dictionary<string, IReadOnlyList<FilaXlsx>>(StringComparer.OrdinalIgnoreCase);
        foreach (var hoja in libro.Root!.Element(ns + "sheets")!.Elements(ns + "sheet"))
        {
            var nombre = hoja.Attribute("name")!.Value;
            var destino = destinos[hoja.Attribute(relacionesNs + "id")!.Value].TrimStart('/');
            if (!destino.StartsWith("xl/", StringComparison.Ordinal))
            {
                destino = "xl/" + destino;
            }

            var entrada = zip.GetEntry(destino);
            if (entrada is null)
            {
                continue;
            }

            XDocument hojaXml;
            using (var s = entrada.Open())
            {
                hojaXml = XDocument.Load(s);
            }

            var filas = new List<FilaXlsx>();
            foreach (var fila in hojaXml.Descendants(ns + "row"))
            {
                var celdas = new Dictionary<string, CeldaXlsx>(StringComparer.Ordinal);
                foreach (var celda in fila.Elements(ns + "c"))
                {
                    var referencia = celda.Attribute("r")!.Value;
                    var tipo = celda.Attribute("t")?.Value;
                    var valor = celda.Element(ns + "v")?.Value;
                    string? texto = tipo switch
                    {
                        "s" when valor is not null => compartidas[int.Parse(valor, CultureInfo.InvariantCulture)],
                        "inlineStr" => string.Concat(celda.Descendants(ns + "is").Elements(ns + "t").Select(t => t.Value)),
                        "str" => valor,
                        _ => null
                    };

                    celdas[LetraDeReferencia(referencia)] = new CeldaXlsx(
                        referencia,
                        tipo,
                        celda.Element(ns + "f")?.Value,
                        valor,
                        texto);
                }

                filas.Add(new FilaXlsx(int.Parse(fila.Attribute("r")!.Value, CultureInfo.InvariantCulture), celdas));
            }

            hojas[nombre] = filas;
        }

        return new LibroXlsx(hojas);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Helpers de dominio y de formato
    // ════════════════════════════════════════════════════════════════════════════════════════════

    private static void AfirmarMonto(string etiqueta, decimal esperado, decimal medido)
    {
        if (Math.Abs(esperado - medido) > Centavo)
        {
            Assert.Fail($"{etiqueta}: esperado {esperado} vs medido {medido} (Δ {medido - esperado}).");
        }
    }

    private static decimal Monto(CeldaXlsx celda)
    {
        // Plan 62 (H5, T5.2 — data drift): en una celda de TEXTO el <v> crudo es el índice de la
        // cadena compartida (o el texto de una fórmula `t="str"`), nunca el importe: se parsea el
        // texto RESUELTO. Solo las celdas numéricas se leen del <v>.
        if (EsCeldaDeTexto(celda.Tipo))
        {
            return Peso(celda.Texto)
                ?? throw new InvalidOperationException($"La celda '{celda.Referencia}' no trae un monto legible.");
        }

        if (!string.IsNullOrWhiteSpace(celda.Valor)
            && decimal.TryParse(celda.Valor, NumberStyles.Float, CultureInfo.InvariantCulture, out var numerico))
        {
            return numerico;
        }

        return Peso(celda.Texto) ?? throw new InvalidOperationException($"La celda '{celda.Referencia}' no trae un monto legible.");
    }

    /// <summary>Parsea un importe es-CO (miles con '.', decimales con ',') o un texto con '$'.</summary>
    private static decimal? Peso(string? texto)
    {
        if (string.IsNullOrWhiteSpace(texto))
        {
            return null;
        }

        var limpio = texto.Replace("$", string.Empty, StringComparison.Ordinal)
            .Replace(" ", string.Empty, StringComparison.Ordinal)
            .Trim();
        if (limpio.Length == 0)
        {
            return null;
        }

        var negativo = limpio.StartsWith('(') || limpio.StartsWith('-');
        limpio = limpio.Trim('(', ')', '-');
        limpio = limpio.Replace(".", string.Empty, StringComparison.Ordinal).Replace(',', '.');
        return decimal.TryParse(limpio, NumberStyles.Number, CultureInfo.InvariantCulture, out var valor)
            ? negativo ? -valor : valor
            : null;
    }

    private static bool EsEspeciales(string? encabezado)
    {
        var normalizado = (encabezado ?? string.Empty).Trim();
        return AliasesEspeciales.Any(alias => normalizado.Equals(alias, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IgualEtiqueta(string valor, string esperado) =>
        string.Equals(NormalizarEtiqueta(valor), NormalizarEtiqueta(esperado), StringComparison.OrdinalIgnoreCase);

    /// <summary>Normaliza un rótulo: sin acentos, sin espacios de borde (los libros alternan ambas formas).</summary>
    private static string NormalizarEtiqueta(string texto)
    {
        var descompuesto = (texto ?? string.Empty).Trim().Normalize(NormalizationForm.FormD);
        var limpio = new StringBuilder(descompuesto.Length);
        foreach (var caracter in descompuesto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(caracter) != UnicodeCategory.NonSpacingMark)
            {
                limpio.Append(caracter);
            }
        }

        return limpio.ToString();
    }

    private static string Texto(object? valor) =>
        valor is null ? string.Empty : Convert.ToString(valor, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;

    private static decimal? NumeroONulo(object? valor) => valor switch
    {
        null => null,
        decimal d => d,
        double d => (decimal)d,
        float f => (decimal)f,
        int i => i,
        long l => l,
        short s => s,
        _ => null
    };

    private static string Letra(int indice)
    {
        var texto = new StringBuilder();
        var numero = indice + 1;
        while (numero > 0)
        {
            var resto = (numero - 1) % 26;
            texto.Insert(0, (char)('A' + resto));
            numero = (numero - 1) / 26;
        }

        return texto.ToString();
    }

    private static int IndiceDeLetra(string letra)
    {
        var indice = 0;
        foreach (var caracter in letra)
        {
            indice = (indice * 26) + (char.ToUpperInvariant(caracter) - 'A' + 1);
        }

        return indice - 1;
    }

    private static string LetraDeReferencia(string? referencia) =>
        new(referencia!.Where(char.IsAsciiLetter).ToArray());

    private static string CarpetaTemporal(string prefijo)
    {
        var carpeta = Path.Combine(Path.GetTempPath(), $"{prefijo}-{Environment.ProcessId}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(carpeta);
        return carpeta;
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
