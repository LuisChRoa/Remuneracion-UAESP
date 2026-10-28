using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ExcelDataReader;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// PLAN 61 — H0/F0 (T0 read-only, G-H0-READONLY): medición en bytes del corpus R4 y de sus oráculos
/// (Administrativo + R10) para el defecto "Especiales R4 leído por índice y con <c>P9 = 0m</c>".
///
/// <para>
/// Naturaleza: este fixture NO corrige nada y NO consume código de producción para decidir: mide el
/// disco. Fuentes R4 con <see cref="ExcelDataReader"/> SOLO-LECTURA sobre COPIAS temporales (nunca sobre
/// el archivo del repo, que puede tener Excel abierto); plantilla y administrativos/R10 con BCL puro
/// (<see cref="ZipArchive"/> + <see cref="XDocument"/>). Rutas explícitas: prohibido
/// <c>Insumos.Raiz()</c> y prohibido localizar por prefijo lo que la evidencia cita con nombre exacto.
/// </para>
///
/// <para>
/// Cada número que afirma un test viaja con su derivación (archivo × etiqueta × celda). Los helpers
/// <c>Verificar*</c> son el MISMO método que usa el test de injerto
/// (<see cref="T0_Injerto_ElMetodoDeMedicionDetectaDivergencia"/>): si la expectativa se altera en un
/// centavo, el método falla. Ningún aserto es la identidad de su propia fórmula.
/// </para>
///
/// <para>
/// Foto: 2026-10-28. El inventario textual (T0.4/T0.5/T0.6) y el defecto descrito son evidencia del
/// ANTES: H1/H2/H5 los modifican y este fixture debe actualizarse en el mismo cambio, con causa citada
/// (baja declarada), nunca en silencio.
/// </para>
/// </summary>
public sealed class EspecialesR4T0Plan61Tests
{
    static EspecialesR4T0Plan61Tests()
    {
        // ExcelReaderConfiguration construye un Encoding de la página 1252: sin el proveedor el
        // constructor del reader falla según el orden en que la suite haya tocado ExcelDataReader.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <summary>Tolerancia de comparación de montos: el centavo (las fuentes traen 2 decimales).</summary>
    private const decimal Centavo = 0.005m;

    /// <summary>
    /// Huella semántica de la plantilla v4 MEDIDA en H0 (algoritmo canónico
    /// <see cref="HuellaSemanticaPlantilla.Calcular"/>). Coincide con <see cref="HuellaSemanticaPlantilla.V4"/>.
    /// El plan 61 (§gate_entrada) cita la misma huella con un dígito distinto
    /// (<c>…a3d11a524d5140b8608c55</c>); el valor medido —el que manda— es el de esta constante.
    /// </summary>
    private const string HuellaV4Medida = "a7476ee2e80dfa13c9e39430d1ea4d77f0c52e3ca0a3d11a526d5140b8608c55";

    /// <summary>
    /// Alias admitidos del componente Especiales, en el orden de la tabla ÚNICA de producción
    /// (<c>ResolvedorConciliacionEmpresa.EncabezadosEspeciales</c>, D-ENCABEZADO ratificada). Aquí se
    /// replica SOLO para medir; <see cref="T0_5_SupuestosDeAusencia_InventarioCerrado"/> comprueba que
    /// la tabla productiva sigue siendo una sola.
    /// </summary>
    private static readonly string[] AliasesEspeciales = ["SERVICIO ESPECIALES", "Especiales"];

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental",
        "2-Lime",
        "3-Ciudad Limpia",
        "4-Bogota Limpia",
        "5-Área Limpia"
    ];

    // ── Rutas explícitas de la evidencia (§0 del plan) ───────────────────────────────────────────

    private static string Raiz => RaizRepositorio();

    private static string RutaPlantillaV4 => Path.Combine(
        Raiz, "Docs", "Plantilla estandarizada", "v4", "Formato Remuneracion 2026xx-x Total v4.xlsx");

    private static string RutaR4Septiembre1Ase2 => Path.Combine(
        Raiz, "Docs", "Prueba Septiembre 1", "Insumos", "2-Lime",
        "ReversiónPorComponente_to_date01092026ddMMyyyy_to_date15092026ddMMyyyy___202691615358215.xlsx");

    private static string RutaR4Septiembre2Ase4 => Path.Combine(
        Raiz, "Docs", "Prueba Septiembre-2", "Insumos", "4-Bogota Limpia",
        "ReversiónPorComponente_to_date16092026ddMMyyyy_to_date30092026ddMMyyyy___20261024531443.xlsx");

    private static string RutaR4Septiembre2Ase2 => Path.Combine(
        Raiz, "Docs", "Prueba Septiembre-2", "Insumos", "2-Lime",
        "ReversiónPorComponente_to_date16092026ddMMyyyy_to_date30092026ddMMyyyy___2026101205321620.xlsx");

    private static string RutaAdministrativoSeptiembre1 => Path.Combine(
        Raiz, "Docs", "Prueba Septiembre 1", "Resultados", "Administrativo", "Remuneracion 202609-1 Total.xlsx");

    private static string RutaAdministrativoSeptiembre2 => Path.Combine(
        Raiz, "Docs", "Prueba Septiembre-2", "Resultados", "Administrativo", "Remuneracion 202609-2 Total.xlsx");

    private static string RutaR10Septiembre1 => Path.Combine(
        Raiz, "Docs", "Prueba Septiembre 1", "Insumos", "R10_Remuneracion_2026091.xlsx");

    private static string RutaR10Septiembre2 => Path.Combine(
        Raiz, "Docs", "Prueba Septiembre-2", "Insumos", "R10_Remuneracion_2026092.xlsx");

    /// <summary>
    /// Nombre REAL de la hoja consolidada: el plan la cita con la forma corta <c>CONSOLIDADO</c>
    /// (mapa de celdas histórico del proyecto); en la plantilla v4 y en los dos administrativos la hoja
    /// se llama <c>CONSOLIDADO_TOTAL RECAUDO</c>. Se mide, no se asume.
    /// </summary>
    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";

    private const string HojaR4 = "Reversion Pagos R4";

    // ── T0.1 · 2026092 / ASE4 (Bogotá Limpia) ────────────────────────────────────────────────────

    [Fact]
    public void T0_1_FuenteR4_2026092_Ase4_FilaLetraEncabezadoYValores()
    {
        var medida = MedirR4(RutaR4Septiembre2Ase4);

        VerificarEntero("2026092 ASE4 · fila de encabezados de componente", 4, medida.FilaEncabezados);
        VerificarEntero("2026092 ASE4 · filas con A='Total' y B vacío", 1, medida.FilasCierre);
        VerificarEntero("2026092 ASE4 · fila de cierre (1-based)", 23, medida.FilaCierre);
        VerificarTexto("2026092 ASE4 · encabezado de la columna de Especiales", "SERVICIO ESPECIALES", medida.EncabezadoEspeciales);
        VerificarTexto("2026092 ASE4 · letra física de la columna de Especiales", "P", medida.LetraEspeciales);
        VerificarMonto("2026092 ASE4 · R4!D23 (Total)", -178619709.53m, medida.D);
        VerificarMonto("2026092 ASE4 · R4!P23 (SERVICIO ESPECIALES)", -49900.44m, medida.EspecialesObservado ?? decimal.Zero);
        VerificarMonto("2026092 ASE4 · D23 − P23 (lo que espera la plantilla)", -178569809.09m, medida.D - medida.Especiales);

        // El defecto, medido: con P9 = 0m el leaf entrega D23 tal cual y el Especiales desaparece.
        VerificarMonto("2026092 ASE4 · Δ del caso = |P23|", 49900.44m, Math.Abs(medida.Especiales));

        // La letra Q trae "Otros": el orden físico de la fuente no es el de la plantilla.
        Assert.Equal("Otros", medida.Encabezados.Single(e => e.Letra == "Q").Encabezado);
    }

    [Fact]
    public void T0_1_Administrativo_2026092_CierreD316YTotalD107()
    {
        var libro = LeerLibro(RutaAdministrativoSeptiembre2);

        Assert.Contains(HojaConsolidado, libro.Hojas);
        Assert.DoesNotContain("CONSOLIDADO", libro.Hojas);
        Assert.Contains(HojaR4, libro.Hojas);

        // Cierre del bloque R4 del ASE4: la plantilla resta AMBOS operandos.
        var d316 = Celda(libro, HojaR4, "D316");
        VerificarTexto("2026092 · 'Reversion Pagos R4'!D316 · fórmula", "D239-P239", d316.Formula);
        VerificarMonto("2026092 · 'Reversion Pagos R4'!D316 · valor", -178569809.09m, Monto(d316));
        VerificarMonto("2026092 · 'Reversion Pagos R4'!D239 (fila Total de la fuente)", -178619709.53m, Monto(Celda(libro, HojaR4, "D239")));
        VerificarMonto("2026092 · 'Reversion Pagos R4'!P239 (SERVICIO ESPECIALES)", -49900.44m, Monto(Celda(libro, HojaR4, "P239")));
        VerificarTexto("2026092 · 'Reversion Pagos R4'!A239", "Total", Celda(libro, HojaR4, "A239").Texto);

        // CONSOLIDADO D69: el operando que hoy la app deja en D239 (P9 = 0m).
        var d69 = Celda(libro, HojaConsolidado, "D69");
        VerificarTexto("2026092 · CONSOLIDADO!D69 · fórmula", "'Reversion Pagos R4'!D316", d69.Formula);
        VerificarMonto("2026092 · CONSOLIDADO!D69 · valor", -178569809.09m, Monto(d69));
        VerificarTexto("2026092 · CONSOLIDADO!N69 · fórmula", "'Reversion Pagos R4'!P316", Celda(libro, HojaConsolidado, "N69").Formula);
        VerificarMonto("2026092 · CONSOLIDADO!N69 · valor", -49900.44m, Monto(Celda(libro, HojaConsolidado, "N69")));

        // Total de la fila del ASE4 en el consolidado del período.
        var d107 = Celda(libro, HojaConsolidado, "D107");
        VerificarTexto("2026092 · CONSOLIDADO!D107 · fórmula", "D12+D31+D50+D69+D88", d107.Formula);
        VerificarMonto("2026092 · CONSOLIDADO!D107 · valor", 7675362501.96m, Monto(d107));

        // Ubicación del "Ajuste Recargo Aseo" en la PLANTILLA (letra Q), frente a la fuente que lo trae
        // en P: la letra no es contrato, el encabezado sí.
        VerificarTexto("2026092 · plantilla · encabezado P del bloque R4", "SERVICIO ESPECIALES", Celda(libro, HojaR4, "P2").Texto);
        VerificarTexto("2026092 · plantilla · encabezado Q del bloque R4", "Ajuste Recargo Aseo", Celda(libro, HojaR4, "Q2").Texto);
    }

    [Fact]
    public void T0_1_R10_2026092_DetRetriYDivergenciaDelDefecto()
    {
        var libro = LeerLibro(RutaR10Septiembre2);
        Assert.Contains("DetRetri2026092", libro.Hojas);
        Assert.DoesNotContain("DetRetri", libro.Hojas);

        VerificarTexto("2026092 · DetRetri · B12 (número de ASE)", "4", Celda(libro, "DetRetri2026092", "B12").Texto);
        VerificarMonto(
            "2026092 · DetRetri2026092!D12 (RECAUDO TOTAL del ASE4)",
            7675362502m,
            Peso(Celda(libro, "DetRetri2026092", "D12").Texto) ?? decimal.MinValue);

        // Δ del caso: el administrativo (con Especiales) vs el número que hoy calcula la app.
        var d107 = 7675362501.96m;
        var conDefecto = d107 - 49900.44m;
        VerificarMonto("2026092 · DetRetri calculado con P9 = 0m", 7675312601.52m, conDefecto);
        Assert.Equal(7675312602m, Math.Round(conDefecto, 0, MidpointRounding.AwayFromZero));
        Assert.Equal(7675362502m, Math.Round(d107, 0, MidpointRounding.AwayFromZero));
        Assert.True(Math.Abs(49900.44m) > 0.5m, "El Δ del caso supera cualquier tolerancia de redondeo.");
    }

    // ── T0.2 · 2026091 / ASE2 (Lime) ────────────────────────────────────────────────────────────

    [Fact]
    public void T0_2_FuenteR4_2026091_Ase2_FilaLetraEncabezadoYValores()
    {
        var medida = MedirR4(RutaR4Septiembre1Ase2);

        VerificarEntero("2026091 ASE2 · fila de encabezados de componente", 4, medida.FilaEncabezados);
        VerificarEntero("2026091 ASE2 · filas con A='Total' y B vacío", 1, medida.FilasCierre);
        VerificarEntero("2026091 ASE2 · fila de cierre (1-based)", 11, medida.FilaCierre);
        VerificarTexto("2026091 ASE2 · encabezado de la columna de Especiales", "SERVICIO ESPECIALES", medida.EncabezadoEspeciales);
        VerificarTexto("2026091 ASE2 · letra física de la columna de Especiales", "P", medida.LetraEspeciales);
        VerificarMonto("2026091 ASE2 · R4!D11 (Total)", -8739602.37m, medida.D);
        VerificarMonto("2026091 ASE2 · R4!P11 (SERVICIO ESPECIALES)", -26500.00m, medida.EspecialesObservado ?? decimal.Zero);
        VerificarMonto("2026091 ASE2 · D11 − P11", -8713102.37m, medida.D - medida.Especiales);
        VerificarMonto("2026091 ASE2 · Δ del caso = |P11|", 26500.00m, Math.Abs(medida.Especiales));
    }

    [Fact]
    public void T0_2_Administrativo_2026091_CierreD160YTotalD105()
    {
        var libro = LeerLibro(RutaAdministrativoSeptiembre1);

        Assert.Contains(HojaConsolidado, libro.Hojas);
        Assert.Contains(HojaR4, libro.Hojas);

        var d160 = Celda(libro, HojaR4, "D160");
        VerificarTexto("2026091 · 'Reversion Pagos R4'!D160 · fórmula", "D97-P97", d160.Formula);
        VerificarMonto("2026091 · 'Reversion Pagos R4'!D160 · valor", -8713102.37m, Monto(d160));
        VerificarMonto("2026091 · 'Reversion Pagos R4'!D97 (fila Total de la fuente)", -8739602.37m, Monto(Celda(libro, HojaR4, "D97")));
        VerificarMonto("2026091 · 'Reversion Pagos R4'!P97 (SERVICIO ESPECIALES)", -26500.00m, Monto(Celda(libro, HojaR4, "P97")));
        VerificarMonto("2026091 · 'Reversion Pagos R4'!P160", -26500.00m, Monto(Celda(libro, HojaR4, "P160")));

        var d67 = Celda(libro, HojaConsolidado, "D67");
        VerificarTexto("2026091 · CONSOLIDADO!D67 · fórmula", "'Reversion Pagos R4'!D160", d67.Formula);
        VerificarMonto("2026091 · CONSOLIDADO!D67 · valor", -8713102.37m, Monto(d67));
        VerificarMonto("2026091 · CONSOLIDADO!N67 · valor", -26500.00m, Monto(Celda(libro, HojaConsolidado, "N67")));

        var d105 = Celda(libro, HojaConsolidado, "D105");
        VerificarTexto("2026091 · CONSOLIDADO!D105 · fórmula", "D10+D29+D48+D67+D86", d105.Formula);
        VerificarMonto("2026091 · CONSOLIDADO!D105 · valor", 14147357755.39m, Monto(d105));
    }

    [Fact]
    public void T0_2_R10_2026091_DetRetriYDivergenciaDelDefecto()
    {
        var libro = LeerLibro(RutaR10Septiembre1);
        Assert.Contains("DetRetri2026091", libro.Hojas);
        Assert.DoesNotContain("DetRetri", libro.Hojas);

        VerificarTexto("2026091 · DetRetri · B10 (número de ASE)", "2", Celda(libro, "DetRetri2026091", "B10").Texto);
        VerificarMonto(
            "2026091 · DetRetri2026091!D10 (RECAUDO TOTAL del ASE2)",
            14147357755m,
            Peso(Celda(libro, "DetRetri2026091", "D10").Texto) ?? decimal.MinValue);

        var d105 = 14147357755.39m;
        var conDefecto = d105 - 26500.00m;
        VerificarMonto("2026091 · DetRetri calculado con P9 = 0m", 14147331255.39m, conDefecto);
        Assert.Equal(14147331255m, Math.Round(conDefecto, 0, MidpointRounding.AwayFromZero));
    }

    // ── T0.3 · Encabezado exacto, normalización y columna intermitente ───────────────────────────

    [Fact]
    public void T0_3_EncabezadoExactoYNormalizacionAplicable()
    {
        // El encabezado medido es exactamente el alias canónico: sin espacios de borde ni variantes.
        Assert.Equal("SERVICIO ESPECIALES", MedirR4(RutaR4Septiembre2Ase4).EncabezadoEspeciales);
        Assert.Equal("SERVICIO ESPECIALES", MedirR4(RutaR4Septiembre1Ase2).EncabezadoEspeciales);

        // Normalización aplicable (D-H1): Trim + OrdinalIgnoreCase sobre la tabla única de alias.
        Assert.True(EsEspeciales("SERVICIO ESPECIALES"));
        Assert.True(EsEspeciales("  servicio especiales  "));
        Assert.True(EsEspeciales("Especiales"));
        Assert.True(EsEspeciales("especiales"));
        Assert.False(EsEspeciales("Ajuste Recargo Aseo"));
        Assert.False(EsEspeciales("Otros"));
        Assert.False(EsEspeciales(string.Empty));

        // La tabla de alias es UNA sola: vive en el resolvedor por empresa y nadie más la re-declara.
        var resolvedor = LeerFuente("Remuneracion.Infrastructure", "Excel", "ResolvedorConciliacionEmpresa.cs");
        VerificarEntero(
            "tabla única de alias de Especiales · declaraciones en el resolvedor",
            1,
            Contar(resolvedor, "[\"SERVICIO ESPECIALES\", \"Especiales\"]"));
    }

    [Fact]
    public void T0_3_Ase2_2026092_LaLetraPTraeAjusteRecargoAseoYNoEspeciales()
    {
        var medida = MedirR4(RutaR4Septiembre2Ase2);

        VerificarEntero("2026092 ASE2 · fila de encabezados", 4, medida.FilaEncabezados);
        Assert.Null(medida.LetraEspeciales);
        Assert.Null(medida.EncabezadoEspeciales);
        VerificarMonto("2026092 ASE2 · Especiales por encabezado (ausente ⇒ 0)", decimal.Zero, medida.Especiales);
        VerificarTexto("2026092 ASE2 · encabezado físico en la letra P", "Ajuste Recargo Aseo", medida.Encabezados.Single(e => e.Letra == "P").Encabezado);
        VerificarTexto("2026092 ASE2 · encabezado físico en la letra Q", "Otros", medida.Encabezados.Single(e => e.Letra == "Q").Encabezado);

        // Lo que un lector por ÍNDICE inyectaría en Especiales: el valor de P de la fila de cierre.
        VerificarMonto("2026092 ASE2 · P22 (Ajuste Recargo Aseo, fila de cierre)", -10.16m, medida.ValoresCierre["P"] ?? decimal.Zero);

        // La plantilla ubica "Ajuste Recargo Aseo" en Q: la fuente lo trae corrido una letra a la izquierda.
        Assert.Equal("Ajuste Recargo Aseo", Celda(LeerLibro(RutaAdministrativoSeptiembre2), HojaR4, "Q2").Texto);

        // ASE5 de 2026092 repite el patrón (P = Ajuste Recargo Aseo, sin Q): no es un caso del ASE2.
        var ase5 = MedirR4(BuscarR4EnCarpeta(Path.Combine(Raiz, "Docs", "Prueba Septiembre-2", "Insumos", CarpetasAse[4])));
        Assert.Null(ase5.LetraEspeciales);
        VerificarTexto("2026092 ASE5 · encabezado físico en la letra P", "Ajuste Recargo Aseo", ase5.Encabezados.Single(e => e.Letra == "P").Encabezado);
    }

    [Fact]
    public void T0_3_PresenciaAusenciaPorAseEnAmbosPeriodos()
    {
        var presencia = new (string Periodo, string Carpeta, int Ase, bool TieneEspeciales)[]
        {
            ("2026091", "Prueba Septiembre 1", 1, false),
            ("2026091", "Prueba Septiembre 1", 2, true),
            ("2026091", "Prueba Septiembre 1", 3, false),
            ("2026091", "Prueba Septiembre 1", 4, false),
            ("2026091", "Prueba Septiembre 1", 5, false),
            ("2026092", "Prueba Septiembre-2", 1, false),
            ("2026092", "Prueba Septiembre-2", 2, false),
            ("2026092", "Prueba Septiembre-2", 3, false),
            ("2026092", "Prueba Septiembre-2", 4, true),
            ("2026092", "Prueba Septiembre-2", 5, false)
        };

        foreach (var caso in presencia)
        {
            var carpeta = Path.Combine(Raiz, "Docs", caso.Carpeta, "Insumos", CarpetasAse[caso.Ase - 1]);
            var medida = MedirR4(BuscarR4EnCarpeta(carpeta));
            Assert.Equal(caso.TieneEspeciales, medida.LetraEspeciales is not null);
            if (caso.TieneEspeciales)
            {
                VerificarTexto($"{caso.Periodo} ASE{caso.Ase} · letra del encabezado", "P", medida.LetraEspeciales);
            }
        }

        // Conteo cerrado por período: exactamente un ASE por quincena declara el encabezado.
        var porPeriodo = presencia
            .Where(p => p.TieneEspeciales)
            .GroupBy(p => p.Periodo)
            .ToDictionary(g => g.Key, g => g.Count());
        Assert.Equal(1, porPeriodo["2026091"]);
        Assert.Equal(1, porPeriodo["2026092"]);
        Assert.Equal(2, presencia.Count(p => p.TieneEspeciales));
    }

    // ── T0.4 · Corpus con conteo cerrado + inventario de gates respecto de la escritura ──────────

    [Fact]
    public void T0_4_CorpusR4_ConteoCerrado()
    {
        var corpus = EnumerarCorpusR4();
        VerificarEntero("corpus R4 · archivos 'ReversiónPorComponente*' bajo Docs", 30, corpus.Count);

        // Jerarquía del corpus: Docs/<Prueba X>/Insumos/<ASE>/archivo.xlsx
        var porPeriodo = corpus
            .GroupBy(ruta => PeriodoDeLaRuta(ruta))
            .ToDictionary(g => g.Key, g => g.Count());
        Assert.Equal(6, porPeriodo.Count);
        Assert.All(porPeriodo.Values, cantidad => Assert.Equal(5, cantidad));

        var conEncabezado = new List<string>();
        var conValor = new List<string>();
        foreach (var ruta in corpus)
        {
            var medida = MedirR4(ruta);
            if (medida.LetraEspeciales is null)
            {
                continue;
            }

            conEncabezado.Add(ruta);
            if ((medida.EspecialesObservado ?? decimal.Zero) != decimal.Zero)
            {
                conValor.Add(ruta);
            }
        }

        VerificarEntero("corpus R4 · archivos con encabezado de Especiales", 2, conEncabezado.Count);
        VerificarEntero("corpus R4 · archivos con valor distinto de cero en Especiales", 2, conValor.Count);
        VerificarTexto("corpus R4 · primer archivo con encabezado", RutaRelativa(RutaR4Septiembre1Ase2), RutaRelativa(conEncabezado[0]));
        VerificarTexto("corpus R4 · segundo archivo con encabezado", RutaRelativa(RutaR4Septiembre2Ase4), RutaRelativa(conEncabezado[1]));

        // Familia distinta (Balance de subsidios y contribuciones): se cuenta para no confundir corpus.
        var balance = Directory
            .EnumerateFiles(Path.Combine(Raiz, "Docs"), "*.xlsx", SearchOption.AllDirectories)
            .Count(ruta => Normalizar(Path.GetFileNameWithoutExtension(ruta)).StartsWith("r4-balancesubsidioycontribuciones", StringComparison.Ordinal));
        VerificarEntero("corpus · archivos 'R4-BalanceSubsidioyContribuciones*' bajo Docs", 31, balance);
    }

    [Fact]
    public void T0_4_GatesYDefecto_PosicionRespectoDeLaEscritura()
    {
        var procesador = LeerFuente("Remuneracion.Core", "Services", "ProcesadorPeriodo.cs");
        var validador = procesador.IndexOf("_validador.Validar(resultado, leafs)", StringComparison.Ordinal);
        var r10 = procesador.IndexOf("ValidarDetRetriContraR10(leafs, r10)", StringComparison.Ordinal);
        var escritura = procesador.IndexOf("_workbookLeafWriter.GenerarWorkbook(solicitud.RutaPlantilla, rutaSinValidar", StringComparison.Ordinal);
        var cruzada27 = procesador.IndexOf("clasificacion.HayNoDeclaradas", StringComparison.Ordinal);
        var promocion = procesador.IndexOf("File.Move(rutaSinValidar, solicitud.RutaSalida", StringComparison.Ordinal);

        Assert.True(validador >= 0, "No se encontró el gate multi-ASE en ProcesadorPeriodo.");
        Assert.True(r10 >= 0, "No se encontró el gate DetRetri-vs-R10 en ProcesadorPeriodo.");
        Assert.True(escritura >= 0, "No se encontró la escritura a 'sin validar' en ProcesadorPeriodo.");
        Assert.True(cruzada27 >= 0, "No se encontró el gate 2.7 en ProcesadorPeriodo.");
        Assert.True(promocion >= 0, "No se encontró la promoción por File.Move en ProcesadorPeriodo.");

        // Inventario T0.4 (posición ANTES/DESPUÉS del writer):
        //   ANTES   · Validar (multi-ASE, incluye sigma de visibles de empresa vs bloque)
        //   ANTES   · ValidarDetRetriContraR10 (ambas quincenas)
        //   DESPUÉS · validación cruzada 2.7 (hay divergencia NO declarada => throw)
        //   DESPUÉS · promoción a nombre limpio (solo si nada falló)
        //
        // RE-ANCLADO EN H3 (Plan 61, §2-D-H3-ii / §12-D5): el inventario del ANTES quedó obsoleto por
        // diseño. La política ratificada invierte el orden: se ESCRIBE primero (siempre hay Excel
        // revisable) y los gates de DATOS corren DESPUÉS clasificando como novedad. La intención que
        // este fixture protege es la que NO puede volver atrás: ningún gate de datos puede quedar
        // antes de la escritura, y la promoción al nombre limpio es siempre el último paso.
        Assert.True(escritura < validador, "Plan 61 H3: el gate multi-ASE corre DESPUÉS de la escritura (degrada a novedad).");
        Assert.True(escritura < r10, "Plan 61 H3: el gate DetRetri-vs-R10 corre DESPUÉS de la escritura (degrada a novedad).");
        Assert.True(escritura < cruzada27, "La validación cruzada 2.7 corre DESPUÉS de la escritura (H3 la degrada a novedad).");
        Assert.True(validador < promocion, "Ningún gate de datos puede quedar después de la promoción al nombre limpio.");
        Assert.True(r10 < promocion, "Ningún gate de datos puede quedar después de la promoción al nombre limpio.");
        Assert.True(cruzada27 < promocion, "La promoción al nombre limpio es el último paso.");
        Assert.True(escritura < promocion, "La escritura a 'sin validar' precede a la promoción.");

        // Gate sigma-empresas: mensaje medido y su recorte a Q1 (residuo R-SIGMA-Q2).
        var basico = LeerFuente("Remuneracion.Core", "Services", "ValidadorBasico.cs");
        Assert.Contains("no coincide con el visible de bloque", basico, StringComparison.Ordinal);
        Assert.Contains("if (leaf.Periodo.NumeroQuincena != 1)", basico, StringComparison.Ordinal);

        // Coherencia leaf-vs-agregado: hoy compara dos números que ignoran Especiales (ciega por construcción).
        var coherencia = LeerFuente("Remuneracion.Infrastructure", "Excel", "WorkbookLeafCoherence.cs");
        Assert.Contains("leaf.R4.TotalReversionEsperada", coherencia, StringComparison.Ordinal);
        Assert.Contains("r4.TotalReversiones", coherencia, StringComparison.Ordinal);
        Assert.Contains("consolidado.ReversionR4", coherencia, StringComparison.Ordinal);

        // El defecto, en las dos puntas del mismo número (R-UN-LADO) — RE-ANCLADO EN H1.
        // H0 fotografió el defecto tal como estaba («var p9 = 0m;» en el leaf y el agregado leyendo
        // solo el Total). H1 lo cierra: la aserción conserva la INTENCIÓN (las dos puntas resuelven
        // Especiales por la tabla ÚNICA de encabezados y el agregado sigue leyendo el Total como
        // observación fiel), no el literal que documentaba el defecto.
        var leafReader = LeerFuente("Remuneracion.Infrastructure", "Excel", "ExcelDataReaderWorkbookLeafInputReader.cs");
        VerificarEntero("H1 · 'var p9 = 0m;' retirado del leaf", 0, Contar(leafReader, "var p9 = 0m;"));
        VerificarContieneMetodo(
            leafReader,
            "private static WorkbookLeafInputsR4 MapearR4(",
            "ResolvedorConciliacionEmpresa.IndicesComponentes(filas)");
        var agregado = LeerFuente("Remuneracion.Infrastructure", "Excel", "ExcelDataReaderRecaudoReader.cs");
        VerificarContieneMetodo(agregado, "public ReversionR4 LeerR4(string rutaArchivo)", "CeldaNumero(filaTotal.ElementAtOrDefault(3))");
        VerificarContieneMetodo(
            agregado,
            "public ReversionR4 LeerR4(string rutaArchivo)",
            "ResolvedorConciliacionEmpresa.IndicesComponentes(filas)");
    }

    // ── T0.5 · Supuestos de ausencia + fixtures con valores inventados ───────────────────────────

    [Fact]
    public void T0_5_SupuestosDeAusencia_InventarioCerrado()
    {
        var leer = LeerFuente("Remuneracion.Infrastructure", "Excel", "ExcelDataReaderWorkbookLeafInputReader.cs");
        var resolvedor = LeerFuente("Remuneracion.Infrastructure", "Excel", "ResolvedorConciliacionEmpresa.cs");
        var mapaQ2 = LeerFuente("Remuneracion.Infrastructure", "Excel", "WorkbookLeafCellMapQ2.cs");
        var validador = LeerFuente("Remuneracion.Infrastructure", "Excel", "ResolutorValorValidacion.cs");

        // (a) Supuestos LATENTES del defecto: RETIRADOS en H1 de los dos lectores (el leaf y el
        // resolvedor por empresa ya no declaran la ausencia de Especiales). El mapa Q2 conserva su
        // comentario porque describe POSICIONES de celda, no valores: queda como deuda declarada
        // para H5.3 (auditoría de supuestos).
        VerificarEntero("supuesto latente · 'sin análogo' en el leaf (retirado en H1)", 0, Contar(leer, "sin análogo"));
        VerificarEntero("supuesto latente · 'patrón histórico' en el leaf (retirado en H1)", 0, Contar(leer, "patrón histórico"));
        VerificarEntero("supuesto latente · 'no tiene análogo' en el resolvedor por empresa (retirado en H1)", 0, Contar(resolvedor, "no tiene análogo"));
        VerificarEntero("supuesto latente · 'sin análogo' en el mapa Q2 (pendiente H5.3)", 1, Contar(mapaQ2, "sin análogo"));

        // (b) Ausencias LEGÍTIMAS y documentadas (la columna no existe => 0), cada una con su comentario.
        VerificarEntero("ausencia legítima · R1 'valores[\"SERVICIO ESPECIALES\"] = 0m;'", 1, Contar(leer, "valores[\"SERVICIO ESPECIALES\"] = 0m;"));
        VerificarEntero("ausencia legítima · R2 detalle 'valores[\"Especiales\"] = 0m;'", 1, Contar(leer, "valores[\"Especiales\"] = 0m;"));
        Assert.Contains("Columna 'Especiales' ausente en la fuente", leer, StringComparison.Ordinal);
        VerificarEntero(
            "ausencia legítima · 'celda ausente = 0' en el resolvedor de validaciones (:191 código, :226 advertencia)",
            2,
            Contar(validador, "celda ausente = 0"));

        // Lista CERRADA: tras H1 sólo el mapa Q2 nombra la ausencia de Especiales (deuda declarada
        // para H5.3). Si aparece un cuarto archivo con el supuesto, el candado textual lo delata.
        var conSupuesto = EnumerarFuentesProductivas()
            .Where(ruta =>
            {
                var texto = File.ReadAllText(ruta);
                return Contar(texto, "sin análogo") + Contar(texto, "patrón histórico") + Contar(texto, "no tiene análogo") > 0;
            })
            .Select(Path.GetFileName)
            .OrderBy(nombre => nombre, StringComparer.Ordinal)
            .ToList();
        Assert.Equal(
            new[] { "WorkbookLeafCellMapQ2.cs" },
            conSupuesto);
    }

    [Fact]
    public void T0_5_FixturesConValoresDeFuenteInventados_InventarioCerrado()
    {
        var fixtures = EnumerarFuentesDeTests();
        var conP9 = fixtures.Where(r => Contar(File.ReadAllText(r), "P9 = 0m") > 0).Select(r => Path.GetFileName(r)!).OrderBy(n => n, StringComparer.Ordinal).ToList();
        var conE26 = fixtures.Where(r => Contar(File.ReadAllText(r), "E26 = 0m") > 0).Select(r => Path.GetFileName(r)!).OrderBy(n => n, StringComparer.Ordinal).ToList();
        var conTautologia = fixtures.Where(r => Contar(File.ReadAllText(r), "leaf.R4.D9 - leaf.R4.P9") > 0).Select(r => Path.GetFileName(r)!).OrderBy(n => n, StringComparer.Ordinal).ToList();

        // RE-ANCLADO EN H2.3 (Plan 61 · doctrina H5.4): el inventario del ANTES era 5 fixtures con
        // «P9 = 0m» manual, 4 con «E26 = 0m» y 2 tautologías. H2.3 los retiró:
        //   · AjustesSfTTests y ConciliacionEmpresaSigmaPlan55Tests: fixtures SINTÉTICOS declarados,
        //     con componentes NO nulos (el neto se conserva) — ya no contienen el literal.
        //   · ConciliacionEmpresaTests, ValidadorBasicoTests y ValidadorMultiAseTests: modelan 2026071
        //     (Julio-1), período en el que la fuente R4 NO declara «SERVICIO ESPECIALES» (medido en el
        //     corpus, §14.3). Allí P9 = 0 es AUSENCIA MEDIDA y cada ocurrencia porta su marca.
        //   · Las tautologías de DetalleR2R4Tests se reemplazaron por valores medidos.
        // Si alguien reintroduce un cero manual sin marca, este candado lo delata.
        Assert.Equal(
            new[] { "ConciliacionEmpresaTests.cs", "ValidadorBasicoTests.cs", "ValidadorMultiAseTests.cs" },
            conP9);
        Assert.Empty(conE26);
        Assert.Empty(conTautologia);

        // Las únicas ocurrencias que sobreviven son ausencias medidas y llevan su marca explícita.
        foreach (var archivo in conP9)
        {
            var texto = File.ReadAllText(Path.Combine(DirectorioTests, archivo));
            VerificarEntero($"ausencia medida · marca '[AUSENCIA-MEDIDA]' en {archivo}", 1, Contar(texto, "[AUSENCIA-MEDIDA]"));
        }

        var detalle = File.ReadAllText(Path.Combine(DirectorioTests, "DetalleR2R4Tests.cs"));
        VerificarEntero("tautología · 'leaf.R4.D9 - leaf.R4.P9' en DetalleR2R4Tests (retirada en H2.3)", 0, Contar(detalle, "leaf.R4.D9 - leaf.R4.P9"));
    }

    // ── T0.6 · Tests que asumen ausencia de salida / nombre limpio / nombre del reporte ─────────

    [Fact]
    public void T0_6_TestsQueAsumenAusenciaDeSalidaONombre()
    {
        var procesadorTests = File.ReadAllText(Path.Combine(DirectorioTests, "ProcesadorPeriodoTests.cs"));

        // RE-ANCLADO EN H3.3: el inventario del ANTES decía que el test de coherencia llamaba al
        // writer directo y exigía "no hay archivo". H3 lo re-ancló preservando la intención
        // (no certificar ≠ no escribir): ahora exige "existe (sin validar) y NO existe el nombre
        // limpio", y pasa por el PROCESADOR (el gate de datos vive ahí desde §12-D5).
        Assert.Contains("WriterMultiAse_GateDeCoherenciaRotoEnAse3_EntregaMarcadaSinNombreLimpio", procesadorTests, StringComparison.Ordinal);
        Assert.Contains("Assert.False(File.Exists(salida)", procesadorTests, StringComparison.Ordinal);
        var cuerpo = ExtraerCuerpo(procesadorTests, "public void WriterMultiAse_GateDeCoherenciaRotoEnAse3_EntregaMarcadaSinNombreLimpio()");
        Assert.Contains("NombresSalida.RutaSinValidar(salida)", cuerpo, StringComparison.Ordinal);
        Assert.Contains("ProcesadorPeriodo(", cuerpo, StringComparison.Ordinal);
        Assert.Contains("Assert.True(File.Exists(marcada)", cuerpo, StringComparison.Ordinal);

        // Los que SÍ conocen el marcado (Plan 43 / O-4): el "sin validar" ya se contempla.
        var salida = File.ReadAllText(Path.Combine(DirectorioTests, "SalidaSinValidarO4Tests.cs"));
        Assert.Contains("Assert.False(File.Exists(salida)", salida, StringComparison.Ordinal);
        Assert.Contains("Assert.True(File.Exists(marcada)", salida, StringComparison.Ordinal);
        Assert.Contains("NombresSalida.RutaSinValidar(salida)", salida, StringComparison.Ordinal);

        var anclas = File.ReadAllText(Path.Combine(DirectorioTests, "AnclasDerivadasH2Plan57Tests.cs"));
        Assert.Contains("File.Exists(salida) ? salida : RutaSinValidar(salida)", anclas, StringComparison.Ordinal);

        // VerificadorRecaudo T19: re-anclado en H3.3 a la política nueva (marcado presente + limpio
        // ausente). Ya NO espera excepción del writer: el gate de datos vive en el procesador.
        var verificador = File.ReadAllText(Path.Combine(Raiz, "Herramientas", "VerificadorRecaudo", "Program.cs"));
        Assert.Contains("EjecutarCasoNegativo(\"T19\"", verificador, StringComparison.Ordinal);
        Assert.Contains("!File.Exists(salida)", verificador, StringComparison.Ordinal);
        var cuerpoT19 = ExtraerCuerpo(verificador, "EjecutarCasoNegativo(\"T19\"");
        Assert.Contains("NombresSalida.RutaSinValidar(salida)", cuerpoT19, StringComparison.Ordinal);
        Assert.Contains("File.Exists(marcada)", cuerpoT19, StringComparison.Ordinal);

        // Nombre del reporte de declaradas: DOS tests dependen del literal exacto (decisión D-TXT).
        var declaradasCli = File.ReadAllText(Path.Combine(DirectorioTests, "DivergenciaDeclaradaCliTests.cs"));
        Assert.Contains(" - Divergencias declaradas.txt", declaradasCli, StringComparison.Ordinal);
        var declaradasReporte = File.ReadAllText(Path.Combine(DirectorioTests, "ReporteDivergenciasDeclaradasTests.cs"));
        Assert.Contains("Divergencias declaradas", declaradasReporte, StringComparison.Ordinal);

        // Plan 61 (D-TXT): el nombre del reporte de novedades es ÚNICO y estable.
        var novedades = File.ReadAllText(Path.Combine(Raiz, "Remuneracion.Infrastructure", "Excel", "ReporteNovedadesWriter.cs"));
        Assert.Contains("public const string SufijoNovedades = \" - Novedades de la corrida.txt\";", novedades, StringComparison.Ordinal);
    }

    [Fact]
    public void T0_6_NingunTestAsumeCodigoDeSalidaMenorOIgualASeis()
    {
        var patronRango = new Regex(@"<=\s*6\b|<\s*7\b", RegexOptions.CultureInvariant);
        var conRango = EnumerarFuentesDeTests()
            .Where(ruta => patronRango.IsMatch(File.ReadAllText(ruta)))
            .Select(Path.GetFileName)
            .ToList();
        Assert.Empty(conRango);

        // RE-ANCLADO EN H3/H4: el contrato creció a 0..7 con `NovedadesEntrega = 7`. La intención del
        // candado se conserva íntegra: ningún test puede asumir un techo de código de salida, y cada
        // constante del contrato está declarada y con su valor exacto (0..6 INTACTOS).
        var codigos = LeerFuente("Remuneracion.Core", "Errors", "CodigosSalida.cs");
        VerificarEntero("contrato de salida · constantes declaradas (0..7 tras Plan 61)", 8, Regex.Matches(codigos, @"public const int \w+ = \d+;").Count);
        Assert.Contains("public const int Ok = 0;", codigos, StringComparison.Ordinal);
        Assert.Contains("public const int DivergenciaDeclaradaReconocida = 6;", codigos, StringComparison.Ordinal);
        Assert.Contains("public const int NovedadesEntrega = 7;", codigos, StringComparison.Ordinal);

        // MicroFixesHu17 fija el 4 con una constante exacta, no con un rango: no bloquea el código 7.
        var micro = File.ReadAllText(Path.Combine(DirectorioTests, "MicroFixesHu17Tests.cs"));
        Assert.Contains("Assert.Equal(4, proceso.ExitCode)", micro, StringComparison.Ordinal);
    }

    // ── T0.7 · Huella semántica v4 (gate de entrada) ────────────────────────────────────────────

    [Fact]
    public void T0_7_HuellaSemanticaV4_IntactaYCoincideConLaCanonica()
    {
        Assert.True(File.Exists(RutaPlantillaV4), $"Falta la plantilla v4 vigente: {RutaPlantillaV4}");

        var medida = HuellaSemanticaPlantilla.Calcular(RutaPlantillaV4);
        VerificarHuella("plantilla v4 · huella medida en H0", HuellaV4Medida, medida);
        VerificarHuella("plantilla v4 · constante canónica del helper", HuellaSemanticaPlantilla.V4, medida);

        Assert.Equal(64, medida.Length);
        Assert.StartsWith("a7476ee2", medida, StringComparison.Ordinal);
        Assert.EndsWith("8c55", medida, StringComparison.Ordinal);
    }

    // ── Injerto: el método de medición detecta una divergencia (no es un assert tautológico) ────

    [Fact]
    public void T0_Injerto_ElMetodoDeMedicionDetectaDivergencia()
    {
        var medida = MedirR4(RutaR4Septiembre2Ase4);

        // Monto alterado en un centavo.
        Assert.Throws<MedicionT0Exception>(() => VerificarMonto("injerto · monto", -178619709.54m, medida.D));
        // Letra física alterada.
        Assert.Throws<MedicionT0Exception>(() => VerificarTexto("injerto · letra", "Q", medida.LetraEspeciales));
        // Fila de cierre alterada.
        Assert.Throws<MedicionT0Exception>(() => VerificarEntero("injerto · fila", 24, medida.FilaCierre));
        // Conteo de corpus alterado.
        Assert.Throws<MedicionT0Exception>(() => VerificarEntero("injerto · corpus", 3, 2));
        // Huella alterada en el último dígito.
        var huella = HuellaSemanticaPlantilla.Calcular(RutaPlantillaV4);
        var alterada = huella[..^1] + (huella[^1] == '0' ? '1' : '0');
        Assert.Throws<MedicionT0Exception>(() => VerificarHuella("injerto · huella", alterada, huella));

        // Control positivo: con los valores medidos, los mismos métodos NO fallan.
        VerificarMonto("control · monto", -178619709.53m, medida.D);
        VerificarTexto("control · letra", "P", medida.LetraEspeciales);
        VerificarEntero("control · fila", 23, medida.FilaCierre);
        VerificarHuella("control · huella", huella, HuellaSemanticaPlantilla.Calcular(RutaPlantillaV4));
    }

    // ── Infraestructura de medición ─────────────────────────────────────────────────────────────

    private sealed class MedicionT0Exception(string mensaje) : Exception(mensaje);

    private sealed record EncabezadoFisico(string Letra, string Encabezado);

    private sealed record MedidaR4(
        string Ruta,
        int Filas,
        int FilaEncabezados,
        IReadOnlyList<EncabezadoFisico> Encabezados,
        int FilasCierre,
        int FilaCierre,
        decimal D,
        string? LetraEspeciales,
        string? EncabezadoEspeciales,
        decimal? EspecialesObservado,
        decimal Especiales,
        IReadOnlyDictionary<string, decimal?> ValoresCierre);

    private sealed record CeldaXlsx(string Referencia, string? Formula, string? Valor, string? Texto);

    private sealed record LibroXlsx(
        IReadOnlyList<string> Hojas,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, CeldaXlsx>> Celdas);

    /// <summary>
    /// Mide una fuente R4 con <see cref="ExcelDataReader"/> SOLO-LECTURA sobre una copia temporal:
    /// fila de encabezados, letra física del encabezado de Especiales (por alias, nunca por índice),
    /// fila de cierre (A='Total' con B vacío) y los valores de esa fila por letra.
    /// </summary>
    private static MedidaR4 MedirR4(string ruta)
    {
        Assert.True(File.Exists(ruta), $"Falta la fuente R4 citada por la evidencia: {ruta}");

        var (copia, carpeta) = CopiarACarpetaTemporal(ruta);
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

            var indiceEncabezados = -1;
            for (var i = 0; i < filas.Count; i++)
            {
                var textos = filas[i].Count(valor => !string.IsNullOrWhiteSpace(Texto(valor)));
                if (textos < 5)
                {
                    continue;
                }

                if (filas[i].Any(valor => Texto(valor).Equals("Total", StringComparison.OrdinalIgnoreCase)))
                {
                    indiceEncabezados = i;
                    break;
                }
            }

            Assert.True(indiceEncabezados >= 0, $"La fuente '{Path.GetFileName(ruta)}' no trae fila de encabezados reconocible.");

            var encabezados = new List<EncabezadoFisico>();
            for (var i = 0; i < filas[indiceEncabezados].Length; i++)
            {
                var texto = Texto(filas[indiceEncabezados][i]);
                if (texto.Length > 0)
                {
                    encabezados.Add(new EncabezadoFisico(Letra(i), texto));
                }
            }

            var cierres = new List<int>();
            for (var i = 0; i < filas.Count; i++)
            {
                if (!Texto(filas[i].ElementAtOrDefault(0)).Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (Texto(filas[i].ElementAtOrDefault(1)).Length == 0)
                {
                    cierres.Add(i);
                }
            }

            Assert.True(cierres.Count >= 1, $"La fuente '{Path.GetFileName(ruta)}' no trae fila de cierre (A='Total' con B vacío).");
            var indiceCierre = cierres[0];
            var filaCierre = filas[indiceCierre];
            var valoresCierre = new Dictionary<string, decimal?>(StringComparer.Ordinal);
            for (var i = 0; i < filaCierre.Length; i++)
            {
                valoresCierre[Letra(i)] = NumeroONulo(filaCierre[i]);
            }

            var letraEspeciales = default(string);
            var encabezadoEspeciales = default(string);
            foreach (var encabezado in encabezados)
            {
                if (!EsEspeciales(encabezado.Encabezado))
                {
                    continue;
                }

                letraEspeciales = encabezado.Letra;
                encabezadoEspeciales = encabezado.Encabezado;
                break;
            }

            decimal? especialesObservado = null;
            if (letraEspeciales is not null)
            {
                especialesObservado = NumeroONulo(filaCierre.ElementAtOrDefault(IndiceDeLetra(letraEspeciales)));
            }

            return new MedidaR4(
                ruta,
                filas.Count,
                indiceEncabezados + 1,
                encabezados,
                cierres.Count,
                indiceCierre + 1,
                NumeroONulo(filaCierre.ElementAtOrDefault(3)) ?? decimal.Zero,
                letraEspeciales,
                encabezadoEspeciales,
                especialesObservado,
                especialesObservado ?? decimal.Zero,
                valoresCierre);
        }
        finally
        {
            Directory.Delete(carpeta, recursive: true);
        }
    }

    private static (string Copia, string Carpeta) CopiarACarpetaTemporal(string ruta)
    {
        var carpeta = Path.Combine(Path.GetTempPath(), "p61-t0-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(carpeta);
        var copia = Path.Combine(carpeta, Path.GetFileName(ruta));
        using var origen = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var destino = new FileStream(copia, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        origen.CopyTo(destino);
        return (copia, carpeta);
    }

    /// <summary>Lectura BCL pura de un libro (plantilla, administrativo, R10): no abre Excel.</summary>
    private static LibroXlsx LeerLibro(string ruta)
    {
        Assert.True(File.Exists(ruta), $"Falta el libro citado por la evidencia: {ruta}");

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

        var entradaLibro = zip.GetEntry("xl/workbook.xml") ?? throw new InvalidOperationException($"'{ruta}' no es un paquete xlsx legible.");
        XDocument libro;
        using (var s = entradaLibro.Open())
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

        var hojas = new List<string>();
        var celdas = new Dictionary<string, IReadOnlyDictionary<string, CeldaXlsx>>(StringComparer.OrdinalIgnoreCase);
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

            var porReferencia = new Dictionary<string, CeldaXlsx>(StringComparer.OrdinalIgnoreCase);
            foreach (var celda in hojaXml.Descendants(ns + "c"))
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
                porReferencia[referencia] = new CeldaXlsx(referencia, celda.Element(ns + "f")?.Value, valor, texto);
            }

            hojas.Add(nombre);
            celdas[nombre] = porReferencia;
        }

        return new LibroXlsx(hojas, celdas);
    }

    private static CeldaXlsx Celda(LibroXlsx libro, string hoja, string referencia)
    {
        Assert.True(libro.Celdas.TryGetValue(hoja, out var celdas), $"El libro no trae la hoja '{hoja}'.");
        Assert.True(celdas!.TryGetValue(referencia, out var celda), $"La hoja '{hoja}' no trae la celda '{referencia}'.");
        return celda!;
    }

    private static decimal Monto(CeldaXlsx celda)
    {
        if (!string.IsNullOrWhiteSpace(celda.Valor) &&
            decimal.TryParse(celda.Valor, NumberStyles.Float, CultureInfo.InvariantCulture, out var numerico))
        {
            return numerico;
        }

        return Peso(celda.Texto) ?? decimal.MinValue;
    }

    /// <summary>Parsea un importe en formato es-CO (miles con '.', decimales con ',') o un texto con '$'.</summary>
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

    private static string Texto(object? valor) =>
        valor is null ? string.Empty : Convert.ToString(valor, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;

    private static decimal? NumeroONulo(object? valor)
    {
        return valor switch
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
    }

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

    private static IReadOnlyList<string> EnumerarCorpusR4() =>
        Directory
            .EnumerateFiles(Path.Combine(Raiz, "Docs"), "*.xlsx", SearchOption.AllDirectories)
            .Where(ruta => Normalizar(Path.GetFileNameWithoutExtension(ruta)).StartsWith("reversionporcomponente", StringComparison.Ordinal))
            .OrderBy(ruta => ruta, StringComparer.Ordinal)
            .ToList();

    /// <summary>Período de una fuente: la carpeta que contiene a <c>Insumos</c> (Docs/Prueba X/Insumos/ASE/…).</summary>
    private static string PeriodoDeLaRuta(string ruta)
    {
        var ase = new DirectoryInfo(Path.GetDirectoryName(ruta)!);
        return ase.Parent?.Parent?.Name ?? string.Empty;
    }

    private static string BuscarR4EnCarpeta(string carpeta)
    {
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta del ASE: {carpeta}");
        var archivos = Directory
            .EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .Where(ruta => Normalizar(Path.GetFileNameWithoutExtension(ruta)).StartsWith("reversionporcomponente", StringComparison.Ordinal))
            .ToList();
        Assert.Single(archivos);
        return archivos[0];
    }

    private static string Normalizar(string texto)
    {
        var descompuesto = texto.Normalize(NormalizationForm.FormD);
        var limpio = new StringBuilder(descompuesto.Length);
        foreach (var caracter in descompuesto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(caracter) != UnicodeCategory.NonSpacingMark)
            {
                limpio.Append(char.ToLowerInvariant(caracter));
            }
        }

        return limpio.ToString();
    }

    private static string DirectorioTests => Path.Combine(Raiz, "Remuneracion.IntegrationTests");

    private static string LeerFuente(params string[] partes) => File.ReadAllText(Path.Combine([Raiz, .. partes]));

    private static IReadOnlyList<string> EnumerarFuentesProductivas() =>
        EnumerarFuentes(
            Path.Combine(Raiz, "Remuneracion.Core"),
            Path.Combine(Raiz, "Remuneracion.Infrastructure"),
            Path.Combine(Raiz, "Remuneracion.Cli"),
            Path.Combine(Raiz, "Remuneracion.WinForms"));

    /// <summary>
    /// Fuentes de test del proyecto, EXCLUYENDO este fixture: si no, sus propias expectativas
    /// textuales se contarían como hallazgos (el inventario mide el resto de la suite).
    /// </summary>
    private static IReadOnlyList<string> EnumerarFuentesDeTests() =>
        EnumerarFuentes(DirectorioTests)
            .Where(ruta => !string.Equals(Path.GetFileName(ruta), "EspecialesR4T0Plan61Tests.cs", StringComparison.Ordinal))
            .ToList();

    private static IReadOnlyList<string> EnumerarFuentes(params string[] raices) =>
        raices
            .SelectMany(raiz => Directory.EnumerateFiles(raiz, "*.cs", SearchOption.AllDirectories))
            .Where(ruta => !ruta.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase)
                && !ruta.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .OrderBy(ruta => ruta, StringComparer.Ordinal)
            .ToList();

    private static string RutaRelativa(string ruta)
    {
        var indice = ruta.IndexOf("Docs", StringComparison.Ordinal);
        return indice < 0 ? ruta : ruta[indice..].Replace('\\', '/');
    }

    private static int Contar(string texto, string aguja)
    {
        var total = 0;
        var indice = 0;
        while ((indice = texto.IndexOf(aguja, indice, StringComparison.Ordinal)) >= 0)
        {
            total++;
            indice += aguja.Length;
        }

        return total;
    }

    private static string ExtraerCuerpo(string texto, string firma)
    {
        var inicio = texto.IndexOf(firma, StringComparison.Ordinal);
        Assert.True(inicio >= 0, $"No se encontró la firma '{firma}'.");
        var siguiente = texto.IndexOf("[Fact]", inicio + firma.Length, StringComparison.Ordinal);
        return siguiente < 0 ? texto[inicio..] : texto[inicio..siguiente];
    }

    private static void VerificarContieneMetodo(string texto, string firma, string contenido)
    {
        var cuerpo = ExtraerCuerpo(texto, firma);
        Assert.Contains(contenido, cuerpo, StringComparison.Ordinal);
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

    // ── Helpers de verificación (los usa también el injerto) ───────────────────────────────────

    private static void VerificarMonto(string etiqueta, decimal esperado, decimal medido)
    {
        if (Math.Abs(esperado - medido) > Centavo)
        {
            throw new MedicionT0Exception($"{etiqueta}: esperado {esperado} vs medido {medido} (delta {medido - esperado}).");
        }
    }

    private static void VerificarTexto(string etiqueta, string esperado, string? medido)
    {
        if (!string.Equals(esperado, medido, StringComparison.Ordinal))
        {
            throw new MedicionT0Exception($"{etiqueta}: esperado '{esperado}' vs medido '{medido}'.");
        }
    }

    private static void VerificarEntero(string etiqueta, int esperado, int medido)
    {
        if (esperado != medido)
        {
            throw new MedicionT0Exception($"{etiqueta}: esperado {esperado} vs medido {medido}.");
        }
    }

    private static void VerificarHuella(string etiqueta, string esperada, string medida)
    {
        if (!string.Equals(esperada, medida, StringComparison.Ordinal))
        {
            throw new MedicionT0Exception($"{etiqueta}: esperada {esperada} vs medida {medida}.");
        }
    }
}
