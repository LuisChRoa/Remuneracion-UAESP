using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 60 (H0/T0 — §5, §6 fila T0): evidencia CONGELADA, SOLO LECTURA, de los hechos que el plan
/// cita en §0. No escribe código de producción: abre los archivos en disco, mide y compara.
///
/// <para>Re-mide: (T0.1) la fuente R2 del ASE2 de octubre-1 y sus importes de <c>Deb/Cred</c>;
/// (T0.2) el oráculo administrativo POR ETIQUETA y la prueba de que su geometría no es la de la
/// plantilla vigente (comparar por fila es inválido, §0.6); (T0.3) la salida actual —el canal
/// <c>O</c> sin mover y <c>Deb/Cred-otros</c> en 0— y las fórmulas del cuadro OPORTUNO de la v4;
/// (T0.4) que ningún cierre de canal supera el umbral en los períodos históricos ni en el resto de
/// ASE de octubre-1, y que <c>Reversion Pagos R4</c> no declara <c>Deb/Cred</c>; (T0.5) la huella
/// semántica de la plantilla vigente.</para>
///
/// <para><b>Regla de oro de este fixture:</b> si un número medido difiere del que el plan cita, el
/// test ENROJECE. La evidencia se re-mide; no se «ajusta» el plan en silencio.</para>
/// </summary>
public sealed class ReclasificacionDebCredT0Plan60Tests
{
    private const decimal Tol = 0.5m;
    private const string Hoja = "Rem. Anticipos R2";
    private const string HojaR4 = "Reversion Pagos R4";

    /// <summary>Nombre exacto de la fuente R2 del ASE2 (Lime) de octubre-1 (§0.3).</summary>
    private static string FuenteAse2Octubre => Ruta(
        "Docs", "Prueba Octubre-1", "Insumos", "2-Lime",
        "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___2026101617441360.xlsx");

    /// <summary>Oráculo administrativo de octubre-1 (§0.4).</summary>
    private static string Oraculo2026101 => Ruta(
        "Docs", "Prueba Octubre-1", "Resultados", "Resultado administrativo", "Remuneracion 202610-1 Total.xlsx");

    /// <summary>Salida de la app ANTES de la regla (§0.5): la divergencia declarada que el plan cierra.</summary>
    private static string SalidaActual2026101 => Ruta(
        "Docs", "Prueba Octubre-1", "Resultados", "5", "Remuneración 202610-1 Total.xlsx");

    /// <summary>Plantilla v4 VIGENTE.</summary>
    private static string PlantillaVigente => Insumos.PlantillaOficialV4();

    // ── T0.1 — fuente R2 del ASE2 ───────────────────────────────────────────────────────────────

    /// <summary>
    /// T0.1 (§0.3): la fuente del ASE2 trae el bloque <c>5..57</c> con headers en la fila <c>4</c>; el
    /// canal <c>O</c> vale 54,00 / 6.759.294,81 / 6.759.348,81; <c>OCCIDENTE</c> 6.759.342,66,
    /// <c>Componente</c> y el gran <c>Total</c> 6.759.371,77; los otros nueve cierres de canal quedan
    /// bajo el umbral; NO hay columna <c>Deb/Cred-otros</c> y la etiqueta <c>O</c> aparece dos veces.
    /// </summary>
    [Fact]
    public void T0_1_FuenteAse2Octubre_MideLoQueElPlanCita()
    {
        Assert.True(File.Exists(FuenteAse2Octubre), $"Falta la fuente citada por §0.3: {FuenteAse2Octubre}");

        using var workbook = SpreadsheetDocument.Open(FuenteAse2Octubre, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var filas = FilasDe(workbookPart, "Sheet1");

        // Headers en la fila 4; la columna del par se resuelve por ENCABEZADO, nunca por letra.
        var encabezados = EncabezadosDe(filas[4], workbookPart);
        Assert.Contains("Deb/Cred", encabezados.Keys);
        Assert.DoesNotContain("Deb/Cred-otros", encabezados.Keys);

        var columnaQ = encabezados["Deb/Cred"];
        var columnaR = SiguienteColumna(columnaQ);

        // Estructura del bloque: headers en 4, primera fila de detalle en 5, gran total en 57.
        Assert.Equal("Vlr Servicio", Texto(filas[5], "D", workbookPart));
        Assert.Equal("Total", Texto(filas[57], "A", workbookPart));
        Assert.Equal("Total", Texto(filas[4], "E", workbookPart));

        // Canal O: dos conceptos y su cierre.
        Cerca(54.00m, Numero(filas[26], columnaQ, workbookPart));
        Cerca(6_759_294.81m, Numero(filas[27], columnaQ, workbookPart));
        Cerca(6_759_348.81m, Numero(filas[28], columnaQ, workbookPart));

        // Empresa, sección y gran total.
        Assert.Equal("OCCIDENTE", Texto(filas[32], "B", workbookPart));
        Cerca(6_759_342.66m, Numero(filas[32], columnaQ, workbookPart));
        Assert.Equal("Componente", Texto(filas[33], "A", workbookPart));
        Cerca(6_759_371.77m, Numero(filas[33], columnaQ, workbookPart));
        Cerca(6_759_371.77m, Numero(filas[57], columnaQ, workbookPart));

        // Los nueve cierres de canal del cuadro Componente, con el importe medido.
        var cierres = CierresDeCanal(filas, columnaQ, workbookPart).Where(c => c.Fila <= 33).ToList();
        Assert.Equal(9, cierres.Count);
        var medidos = cierres.ToDictionary(c => c.Canal, c => c.Valor, StringComparer.OrdinalIgnoreCase);
        Cerca(5.17m, medidos["A"]);
        Cerca(31.40m, medidos["E"]);
        Cerca(12.95m, medidos["1"]);
        Cerca(2.44m, medidos["2"]);
        Cerca(4.90m, medidos["3"]);
        Cerca(-27.75m, medidos["5"]);
        Cerca(-21.95m, medidos["H"]);
        Cerca(6_759_348.81m, medidos["O"]);
        Cerca(15.80m, medidos["T"]);

        // Un único cierre supera el umbral en TODO el bloque: el del canal O.
        var sobreUmbral = CierresDeCanal(filas, columnaQ, workbookPart)
            .Where(c => Math.Abs(c.Valor) > ReclasificadorDebCredR2.UmbralReclasificacionDebCred)
            .ToList();
        Assert.Single(sobreUmbral);
        Assert.Equal("O", sobreUmbral[0].Canal);

        // La etiqueta 'O' aparece DOS veces (dos instancias, dos cuadros).
        Assert.Equal(2, CierresDeCanal(filas, columnaQ, workbookPart)
            .Count(c => string.Equals(c.Canal, "O", StringComparison.OrdinalIgnoreCase)));

        // La fuente no trae la columna Deb/Cred-otros: todos los R del bloque están vacíos.
        for (var fila = 5; fila <= 57; fila++)
        {
            Assert.Equal(string.Empty, Texto(filas[fila], columnaR, workbookPart));
        }
    }

    // ── T0.2 — oráculo administrativo, por etiqueta ─────────────────────────────────────────────

    /// <summary>
    /// T0.2 (§0.4/§0.6): el oráculo mueve el bloque del canal <c>O</c> a <c>Deb/Cred-otros</c>
    /// (54,00 / 6.759.294,81 / 6.759.348,81 con <c>Deb/Cred</c> = 0), deja <c>OCCIDENTE</c> en
    /// −6,15 / +6.759.348,81 y <c>Componente</c> y el gran <c>Total</c> en 22,96 / +6.759.348,81.
    /// Su geometría NO es la de la plantilla vigente: comparar por número de fila es inválido.
    /// </summary>
    [Fact]
    public void T0_2_Oraculo_PorEtiqueta_GeometriaDistintaALaVigente()
    {
        var oraculo = BloqueDelAse(Oraculo2026101, Hoja, aseId: 2);

        Cerca(0m, oraculo[Clave("Componente", "", "", "", "Vlr Servicio", "O")].Q);
        Cerca(54.00m, oraculo[Clave("Componente", "", "", "", "Vlr Servicio", "O")].R);
        Cerca(0m, oraculo[Clave("Componente", "", "", "", "Vlr Intereses", "O")].Q);
        Cerca(6_759_294.81m, oraculo[Clave("Componente", "", "", "", "Vlr Intereses", "O")].R);
        Cerca(0m, oraculo[Clave("Componente", "", "", "O", "Total", "O")].Q);
        Cerca(6_759_348.81m, oraculo[Clave("Componente", "", "", "O", "Total", "O")].R);
        Cerca(-6.15m, oraculo[Clave("Componente", "", "OCCIDENTE", "Total", "", null)].Q);
        Cerca(6_759_348.81m, oraculo[Clave("Componente", "", "OCCIDENTE", "Total", "", null)].R);
        Cerca(22.96m, oraculo[Clave("Componente", "Componente", "Total", "", "", null)].Q);
        Cerca(6_759_348.81m, oraculo[Clave("Componente", "Componente", "Total", "", "", null)].R);
        Cerca(22.96m, oraculo[Clave("", "Total", "", "", "", null)].Q);
        Cerca(6_759_348.81m, oraculo[Clave("", "Total", "", "", "", null)].R);

        // Las dos filas del cuadro OPORTUNO del ASE2 son FÓRMULA en la v4 (§0.2): el oráculo las
        // publica con su valor evaluado; se citan por etiqueta, nunca por coordenada.
        var oportuno = CuadroOportunoDelAse(Oraculo2026101, Hoja, aseId: 2);
        Cerca(-6.15m, oportuno["OCCIDENTE"].Q);
        Cerca(6_759_348.81m, oportuno["OCCIDENTE"].R);
        Cerca(22.96m, oportuno["TOTAL"].Q);
        Cerca(6_759_348.81m, oportuno["TOTAL"].R);

        // §0.6: la fila de 'OCCIDENTE' del bloque NO coincide con la de la v4 ⇒ comparar por fila
        // sería un quemado de coordenada. El T0 lo fija como hecho medido.
        Assert.NotEqual(
            FilaDeEtiqueta(Oraculo2026101, Hoja, aseId: 2, Clave("Componente", "", "OCCIDENTE", "Total", "", null)),
            FilaDeEtiqueta(PlantillaVigente, Hoja, aseId: 2, Clave("Componente", "", "OCCIDENTE", "Total", "", null)));
    }

    // ── T0.3 — salida actual de la app + fórmulas de la v4 ──────────────────────────────────────

    /// <summary>
    /// T0.3 (§0.5/§0.2): la salida actual deja los tres importes del canal <c>O</c> en
    /// <c>Deb/Cred</c> y <c>Deb/Cred-otros</c> en 0 en todo el bloque; y la v4 declara fórmulas en
    /// EXACTAMENTE las celdas del cuadro OPORTUNO que fija §0.2.
    /// </summary>
    [Fact]
    public void T0_3_SalidaActual_SinMover_Y_FormulasDelOportunoEnLaVigente()
    {
        var actual = BloqueDelAse(SalidaActual2026101, Hoja, aseId: 2);

        Cerca(54.00m, actual[Clave("Componente", "", "", "", "Vlr Servicio", "O")].Q);
        Cerca(6_759_294.81m, actual[Clave("Componente", "", "", "", "Vlr Intereses", "O")].Q);
        Cerca(6_759_348.81m, actual[Clave("Componente", "", "", "O", "Total", "O")].Q);
        Cerca(6_759_342.66m, actual[Clave("Componente", "", "OCCIDENTE", "Total", "", null)].Q);
        Cerca(6_759_371.77m, actual[Clave("Componente", "Componente", "Total", "", "", null)].Q);
        Assert.All(actual.Values, valor => Cerca(0m, valor.R));

        // §0.2: la lista congelada de celdas Q/R con fórmula en la plantilla VIGENTE (18 filas × Q,R).
        // Es la evidencia que impide escribir el cuadro OPORTUNO: se recalcula solo.
        string[] esperadas =
        [
            "Q57", "R57", "Q63", "R63", "Q69", "R69",
            "Q178", "R178", "Q181", "R181", "Q184", "R184", "Q190", "R190",
            "Q299", "R299", "Q305", "R305", "Q311", "R311",
            "Q423", "R423", "Q426", "R426", "Q429", "R429", "Q435", "R435",
            "Q510", "R510", "Q516", "R516", "Q519", "R519", "Q522", "R522",
        ];

        var conFormula = CeldasDelParConFormula(PlantillaVigente, Hoja);
        Assert.Equal(
            esperadas.OrderBy(r => r, StringComparer.Ordinal).ToArray(),
            conFormula.OrderBy(r => r, StringComparer.Ordinal).ToArray());

        // Las celdas del par dentro del bloque de detalle son VALORES literales (no fórmulas).
        Assert.Empty(CeldasDelParDelBloqueConFormula(PlantillaVigente, Hoja, aseId: 2));
    }

    // ── T0.4 — regresión histórica y alcance R4 ────────────────────────────────────────────────

    /// <summary>
    /// T0.4 (§0.7/D3): en julio-1, julio-2 y agosto-2 ningún cierre de canal supera el umbral (la
    /// regla es no-op ⇒ <c>Deb/Cred-otros</c> queda en 0); en octubre-1 sólo dispara el canal <c>O</c>
    /// del ASE2 (ASE3 llega a 2.232,90; ASE1/4/5 no traen <c>Deb/Cred</c>); y
    /// <c>Reversion Pagos R4</c> no declara los puertos del par.
    /// </summary>
    [Fact]
    public void T0_4_HistoricoSinDisparo_Y_R4SinDebCred()
    {
        var periodos = new (string Codigo, string[] Carpeta)[]
        {
            ("2026071", ["Docs", "Prueba Julio-1", "Insumos"]),
            ("2026072", ["Docs", "Prueba Julio-2", "Insumos"]),
            ("2026082", ["Docs", "Prueba Agosto-2", "Insumos"]),
        };

        var lector = new ExcelDataReaderWorkbookLeafInputReader();
        foreach (var (codigo, carpeta) in periodos)
        {
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var r2 = BuscarFuente(Ruta(carpeta), aseId, "RerpoteDetalleSaldosaFavor");
                var maximo = MaximoCierreDeCanal(lector.LeerDetalleR2(Insumos.Ase(aseId), r2));
                Assert.True(
                    Math.Abs(maximo) <= ReclasificadorDebCredR2.UmbralReclasificacionDebCred,
                    $"{codigo} · ASE {aseId}: un cierre de canal trae |Deb/Cred| = {maximo} (> umbral); §0.7 mediría otra cosa.");
            }
        }

        var maximos = new Dictionary<int, decimal>();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var r2 = BuscarFuente(Ruta("Docs", "Prueba Octubre-1", "Insumos"), aseId, "RerpoteDetalleSaldosaFavor");
            maximos[aseId] = MaximoCierreDeCanal(lector.LeerDetalleR2(Insumos.Ase(aseId), r2));
        }

        Cerca(6_759_348.81m, maximos[2]);
        Cerca(2_232.90m, maximos[3]);

        // §0.3 dice que ASE1/ASE4/ASE5 «no traen valores en Deb/Cred (columna presente, celdas
        // vacías)». MEDIDO en el disco: SÍ traen valores, todos MUY por debajo del umbral
        // (160,05 / 98,04 / 124,42), de modo que la conclusión del plan —en octubre-1 la regla
        // dispara UNA sola vez, ASE2 · canal O— se conserva. La diferencia de redacción queda
        // declarada en Docs/T0-Plan60-Reclasificacion-DebCred.md, no silenciada.
        Cerca(160.05m, maximos[1]);
        Cerca(98.04m, maximos[4]);
        Cerca(124.42m, maximos[5]);
        Assert.All(
            new[] { 1, 3, 4, 5 },
            aseId => Assert.True(
                Math.Abs(maximos[aseId]) <= ReclasificadorDebCredR2.UmbralReclasificacionDebCred,
                $"ASE {aseId}: un cierre de canal trae |Deb/Cred| = {maximos[aseId]} (> umbral): la regla dispararía donde §0.3 dice que no."));

        // D3: R4 no declara los puertos del par (su última columna de valor es 'Otros').
        var encabezadosR4 = EncabezadosDelBloque(PlantillaVigente, HojaR4, aseId: 1);
        Assert.DoesNotContain("Deb/Cred", encabezadosR4);
        Assert.DoesNotContain("Deb/Cred-otros", encabezadosR4);

        // Y los manuales históricos publican Deb/Cred-otros en 0 dentro de los bloques de detalle.
        foreach (var manual in new[]
                 {
                     Ruta("Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx"),
                     Ruta("Docs", "Prueba Agosto-2", "Resultado", "Resultado Manual por el administrativo", "Remuneracion 202608-2 Total_7721.xlsx"),
                 })
        {
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var bloque = BloqueDelAse(manual, Hoja, aseId);
                Assert.All(
                    bloque.Values,
                    valor => Assert.True(
                        Math.Abs(valor.R) <= Tol,
                        $"{Path.GetFileName(manual)} · ASE {aseId}: Deb/Cred-otros = {valor.R} (debía ser 0, §0.7)."));
            }
        }
    }

    // ── T0.5 — huella semántica de la plantilla vigente ────────────────────────────────────────

    /// <summary>T0.5 (G-HUELLA): la plantilla vigente conserva la huella semántica del plan.</summary>
    [Fact]
    public void T0_5_HuellaSemanticaDeLaPlantillaVigente()
    {
        Assert.Equal(HuellaSemanticaPlantilla.V4, HuellaSemanticaPlantilla.Calcular(PlantillaVigente));
    }

    // ── medición: bloque por etiqueta ───────────────────────────────────────────────────────────

    /// <summary>
    /// Bloque de detalle del ASE indexado por <c>sección|A|B|C|D|canal</c>. La sección es la
    /// etiqueta <c>A</c> del cierre de sección que encierra la fila; el canal es el rótulo de grupo
    /// que la cierra (<see cref="CanalEfectivoDetalleR2R4"/>). Nunca se usa el número de fila.
    /// </summary>
    private static Dictionary<string, (decimal Q, decimal R)> BloqueDelAse(string ruta, string nombreHoja, int aseId)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = Delimitar(workbookPart, nombreHoja, aseId);
        var filas = FilasDe(workbookPart, nombreHoja);
        var columnas = EncabezadosDe(filas[bloque.EncabezadoColumnas], workbookPart);
        Assert.Contains("Deb/Cred", columnas.Keys);
        Assert.Contains("Deb/Cred-otros", columnas.Keys);

        var rotulos = new List<CanalEfectivoDetalleR2R4.Rotulo>();
        for (var fila = bloque.EncabezadoColumnas + 1; fila <= bloque.GranTotal; fila++)
        {
            rotulos.Add(new CanalEfectivoDetalleR2R4.Rotulo(
                Texto(filas[fila], "A", workbookPart),
                Texto(filas[fila], "B", workbookPart),
                Texto(filas[fila], "C", workbookPart),
                Texto(filas[fila], "D", workbookPart)));
        }

        var secciones = Secciones(rotulos);
        var mapa = new Dictionary<string, (decimal Q, decimal R)>(StringComparer.Ordinal);
        for (var i = 0; i < rotulos.Count; i++)
        {
            var canal = CanalEfectivoDetalleR2R4.CanalEfectivo(
                rotulos,
                i,
                WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR2,
                WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR2);
            var fila = bloque.EncabezadoColumnas + 1 + i;
            mapa[Clave(secciones[i], rotulos[i].A, rotulos[i].B, rotulos[i].C, rotulos[i].D, canal)] =
                (Numero(filas[fila], columnas["Deb/Cred"], workbookPart),
                 Numero(filas[fila], columnas["Deb/Cred-otros"], workbookPart));
        }

        return mapa;
    }

    /// <summary>Cuadro OPORTUNO del ASE (entre el gran total de su bloque y la cabecera siguiente), por empresa.</summary>
    private static Dictionary<string, (decimal Q, decimal R)> CuadroOportunoDelAse(
        string ruta,
        string nombreHoja,
        int aseId)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = Delimitar(workbookPart, nombreHoja, aseId);
        var filas = FilasDe(workbookPart, nombreHoja);
        var columnas = EncabezadosDe(filas[bloque.EncabezadoColumnas], workbookPart);

        var cuadro = new Dictionary<string, (decimal Q, decimal R)>(StringComparer.OrdinalIgnoreCase);
        for (var fila = bloque.GranTotal + 1; fila <= bloque.GranTotal + 60 && filas.ContainsKey(fila); fila++)
        {
            var a = Texto(filas[fila], "A", workbookPart);
            var b = Texto(filas[fila], "B", workbookPart);
            var c = Texto(filas[fila], "C", workbookPart);
            var d = Texto(filas[fila], "D", workbookPart);

            if (a.Length == 0 && b.Length > 0 && c.Length == 0 && d.Length == 0)
            {
                break; // cabecera del bloque siguiente
            }

            if (b.Length > 0 && string.Equals(c, "OPORTUNO", StringComparison.OrdinalIgnoreCase)
                && string.Equals(d, "Total", StringComparison.OrdinalIgnoreCase))
            {
                cuadro[b] = (Numero(filas[fila], columnas["Deb/Cred"], workbookPart),
                             Numero(filas[fila], columnas["Deb/Cred-otros"], workbookPart));
            }
        }

        return cuadro;
    }

    /// <summary>Número de fila (relativo a la hoja) de la etiqueta indicada. Solo para el hecho §0.6.</summary>
    private static int FilaDeEtiqueta(string ruta, string nombreHoja, int aseId, string clave)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = Delimitar(workbookPart, nombreHoja, aseId);
        var filas = FilasDe(workbookPart, nombreHoja);

        var rotulos = new List<CanalEfectivoDetalleR2R4.Rotulo>();
        for (var fila = bloque.EncabezadoColumnas + 1; fila <= bloque.GranTotal; fila++)
        {
            rotulos.Add(new CanalEfectivoDetalleR2R4.Rotulo(
                Texto(filas[fila], "A", workbookPart),
                Texto(filas[fila], "B", workbookPart),
                Texto(filas[fila], "C", workbookPart),
                Texto(filas[fila], "D", workbookPart)));
        }

        var secciones = Secciones(rotulos);
        for (var i = 0; i < rotulos.Count; i++)
        {
            var canal = CanalEfectivoDetalleR2R4.CanalEfectivo(
                rotulos,
                i,
                WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR2,
                WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR2);
            if (string.Equals(Clave(secciones[i], rotulos[i].A, rotulos[i].B, rotulos[i].C, rotulos[i].D, canal), clave, StringComparison.Ordinal))
            {
                return bloque.EncabezadoColumnas + 1 + i;
            }
        }

        throw new InvalidOperationException($"La etiqueta '{clave}' no existe en el bloque del ASE {aseId} de '{nombreHoja}'.");
    }

    /// <summary>Comparación monetaria del T0: centavo (los .xlsx guardan dobles; 0,01 es el grano real).</summary>
    private static void Cerca(decimal esperado, decimal medido) =>
        Assert.True(Math.Abs(medido - esperado) <= 0.01m, $"esperado {esperado} · medido {medido} (difieren en más de un centavo).");

    private static IReadOnlyList<string> Secciones(IReadOnlyList<CanalEfectivoDetalleR2R4.Rotulo> rotulos)
    {
        var secciones = new string[rotulos.Count];
        for (var i = 0; i < rotulos.Count; i++)
        {
            secciones[i] = string.Empty;
            for (var j = i; j < rotulos.Count; j++)
            {
                var a = rotulos[j].A.Trim();
                if (a.Length > 0 && string.Equals(rotulos[j].B.Trim(), "Total", StringComparison.OrdinalIgnoreCase))
                {
                    secciones[i] = a;
                    break;
                }
            }
        }

        return secciones;
    }

    private static string Clave(string seccion, string a, string b, string c, string d, string? canal) =>
        $"{seccion}|{a.Trim()}|{b.Trim()}|{c.Trim()}|{d.Trim()}|{canal}";

    private static IReadOnlyList<(int Fila, string Canal, decimal Valor)> CierresDeCanal(
        IReadOnlyDictionary<int, Row> filas,
        string columnaQ,
        WorkbookPart workbookPart)
    {
        var cierres = new List<(int, string, decimal)>();
        foreach (var (numero, fila) in filas.OrderBy(f => f.Key))
        {
            var canal = Texto(fila, "C", workbookPart);
            if (canal.Length == 0)
            {
                continue;
            }

            if (!string.Equals(Texto(fila, "D", workbookPart), "Total", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (Texto(fila, "A", workbookPart).Length > 0 || Texto(fila, "B", workbookPart).Length > 0)
            {
                continue;
            }

            cierres.Add((numero, canal, Numero(fila, columnaQ, workbookPart)));
        }

        return cierres;
    }

    private static decimal MaximoCierreDeCanal(DetalleR2AseInputs detalle) =>
        detalle.Filas
            .Where(f => string.Equals(f.D.Trim(), "Total", StringComparison.OrdinalIgnoreCase) && f.C.Trim().Length > 0)
            .Select(f => Math.Abs(f.Valor(ReclasificadorDebCredR2.ComponenteDebCred) ?? 0m))
            .DefaultIfEmpty(0m)
            .Max();

    private static string BuscarFuente(string carpetaPeriodo, int aseId, string prefijo)
    {
        var carpeta = Directory.GetDirectories(carpetaPeriodo)
            .First(d => Path.GetFileName(d).StartsWith($"{aseId}-", StringComparison.Ordinal));
        return Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .First(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
    }

    private static DelimitadorBloquesHoja.BloqueHoja Delimitar(WorkbookPart workbookPart, string nombreHoja, int aseId) =>
        DelimitadorBloquesHoja.DelimitarBloque(
            workbookPart,
            string.Equals(nombreHoja, HojaR4, StringComparison.OrdinalIgnoreCase) ? DelimitadorBloquesHoja.R4 : DelimitadorBloquesHoja.R2,
            aseId);

    private static IReadOnlyList<string> EncabezadosDelBloque(string ruta, string nombreHoja, int aseId)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = Delimitar(workbookPart, nombreHoja, aseId);
        var filas = FilasDe(workbookPart, nombreHoja);
        return EncabezadosDe(filas[bloque.EncabezadoColumnas], workbookPart).Keys.ToList();
    }

    /// <summary>Celdas del par (Q/R) con fórmula en TODA la hoja, por referencia.</summary>
    private static IReadOnlyList<string> CeldasDelParConFormula(string ruta, string nombreHoja)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = HojaDe(workbookPart, nombreHoja);
        var tabla = worksheet.Elements<SheetData>().FirstOrDefault()
            ?? throw new InvalidOperationException($"La hoja '{nombreHoja}' no tiene SheetData.");

        return tabla.Elements<Row>()
            .SelectMany(f => f.Elements<Cell>())
            .Where(c => c.CellFormula is not null && c.CellReference?.Value is { } r && EsDelPar(ColumnaDe(r)))
            .Select(c => c.CellReference!.Value!)
            .ToList();
    }

    /// <summary>Celdas del par DENTRO del bloque de detalle que traen fórmula (no debe haber ninguna).</summary>
    private static IReadOnlyList<string> CeldasDelParDelBloqueConFormula(string ruta, string nombreHoja, int aseId)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = Delimitar(workbookPart, nombreHoja, aseId);
        var filas = FilasDe(workbookPart, nombreHoja);

        var conFormula = new List<string>();
        for (var fila = bloque.EncabezadoColumnas + 1; fila <= bloque.GranTotal; fila++)
        {
            foreach (var celda in filas[fila].Elements<Cell>())
            {
                if (celda.CellFormula is not null
                    && celda.CellReference?.Value is { } referencia
                    && EsDelPar(ColumnaDe(referencia)))
                {
                    conFormula.Add(referencia);
                }
            }
        }

        return conFormula;
    }

    private static bool EsDelPar(string columna) =>
        string.Equals(columna, "Q", StringComparison.OrdinalIgnoreCase)
        || string.Equals(columna, "R", StringComparison.OrdinalIgnoreCase);

    private static string Ruta(params string[] partes) =>
        Path.Combine([Insumos.Raiz(), .. partes]);

    // ── lectura de bajo nivel ───────────────────────────────────────────────────────────────────

    private static IReadOnlyDictionary<int, Row> FilasDe(WorkbookPart workbookPart, string nombreHoja)
    {
        var worksheet = HojaDe(workbookPart, nombreHoja);
        var tabla = worksheet.Elements<SheetData>().FirstOrDefault()
            ?? throw new InvalidOperationException($"La hoja '{nombreHoja}' no tiene SheetData.");
        return tabla.Elements<Row>()
            .Where(f => f.RowIndex is not null)
            .GroupBy(f => (int)f.RowIndex!.Value)
            .ToDictionary(g => g.Key, g => g.First());
    }

    private static Worksheet HojaDe(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet.");
    }

    private static Dictionary<string, string> EncabezadosDe(Row fila, WorkbookPart workbookPart)
    {
        var encabezados = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var celda in fila.Elements<Cell>())
        {
            var texto = TextoDe(celda, workbookPart).Trim();
            if (texto.Length > 0 && celda.CellReference?.Value is { } referencia)
            {
                encabezados[texto] = ColumnaDe(referencia);
            }
        }

        return encabezados;
    }

    private static string SiguienteColumna(string columna)
    {
        var indice = 0;
        foreach (var letra in columna)
        {
            indice = (indice * 26) + (char.ToUpperInvariant(letra) - 'A' + 1);
        }

        return NombreDeColumna(indice + 1);
    }

    private static string NombreDeColumna(int indice)
    {
        var nombre = string.Empty;
        while (indice > 0)
        {
            var resto = (indice - 1) % 26;
            nombre = (char)('A' + resto) + nombre;
            indice = (indice - 1) / 26;
        }

        return nombre;
    }

    private static string Texto(Row fila, string columna, WorkbookPart workbookPart)
    {
        var referencia = columna + fila.RowIndex!.Value.ToString(CultureInfo.InvariantCulture);
        var celda = fila.Elements<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        return celda is null ? string.Empty : TextoDe(celda, workbookPart).Trim();
    }

    private static string TextoDe(Cell celda, WorkbookPart workbookPart)
    {
        if (celda.DataType?.Value == CellValues.SharedString
            && int.TryParse(celda.CellValue?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice))
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable?
                .Elements<SharedStringItem>().ElementAtOrDefault(indice)?.InnerText ?? string.Empty;
        }

        return celda.CellValue?.InnerText ?? string.Empty;
    }

    private static decimal Numero(Row fila, string columna, WorkbookPart workbookPart)
    {
        var texto = Texto(fila, columna, workbookPart);
        return decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor) ? valor : 0m;
    }

    private static string ColumnaDe(string referencia)
    {
        var fin = 0;
        while (fin < referencia.Length && (char.IsLetter(referencia[fin]) || referencia[fin] == '$'))
        {
            fin++;
        }

        return referencia[..fin].Replace("$", string.Empty, StringComparison.Ordinal);
    }
}
