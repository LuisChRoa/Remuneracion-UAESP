using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (§11-bis) + Plan 60 (H3.1 — §6 I1..I4, G3..G6): el DESGLOSE-DETALLE por componente de
/// <c>Rem. Anticipos R2</c> y <c>Reversion Pagos R4</c> se escribe en AMBAS quincenas, con sus tres
/// niveles —detalle por canal/concepto, totales por canal/empresa y cierre de sección—, y lo escrito
/// se contrasta contra lo que el DOMINIO declara y contra el oráculo administrativo.
///
/// <para><b>Plan 60 — el gate queda CERRADO:</b> hasta este plan la partición
/// <c>Deb/Cred</c> / <c>Deb/Cred-otros</c> se comparaba como PAR porque la reclasificación por canal
/// no existía (el administrativo movía a «otros» el bloque del canal <c>O</c> de ASE2 y la app lo
/// dejaba en <c>Deb/Cred</c>). Con <see cref="ReclasificadorDebCredR2"/> aplicada en el writer, las
/// dos columnas se comparan <b>POR SEPARADO y por ETIQUETA</b> (§0.4/§0.6: nunca por número de fila
/// — la geometría del oráculo es de una plantilla anterior).</para>
///
/// <para>Lo que se prueba, sin escribir una sola coordenada (todo sale de la gramática de cierres de
/// la hoja, de los encabezados que la propia hoja declara y de
/// <see cref="CanalEfectivoDetalleR2R4"/>):</para>
/// <list type="number">
/// <item>I1: octubre-1 · ASE2 · canal <c>O</c> reclasificado — Q y R por separado contra el oráculo.</item>
/// <item>I2: el resto de filas del bloque ASE2 sigue con <c>Deb/Cred-otros</c> en 0 y su
/// <c>Deb/Cred</c> intacto.</item>
/// <item>I3: ASE1/3/4/5 intactos (ningún canal supera el umbral).</item>
/// <item>I4: la traza de auditoría del caso trae exactamente los tres conceptos del canal <c>O</c>
/// con su coordenada completa, y el grupo homónimo del cuadro <c>Subs/Cont</c> no produce ninguno.</item>
/// <item>G3: blast radius — lo único que cambia frente a la fuente es el par Q/R de esos seis
/// registros, y el par <c>Q+R</c> se conserva (CONSOLIDADO/VALIDACION siguen cerrando).</item>
/// <item>G4: el cuadro OPORTUNO sigue siendo fórmula, con el MISMO texto que la plantilla vigente.</item>
/// <item>G5: julio-1, julio-2 y agosto-2 no disparan (la regla es no-op: ninguna celda Q/R cambia).</item>
/// <item>G6: <c>Reversion Pagos R4</c> no tiene puertos <c>Deb/Cred</c> ⇒ la regla no corre allí.</item>
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

    private const string ColumnaDebCred = ReclasificadorDebCredR2.ComponenteDebCred;
    private const string ColumnaDebCredOtros = ReclasificadorDebCredR2.ComponenteDebCredOtros;

    /// <summary>
    /// Etiquetas (sección · firma A|B|C|D · canal efectivo) que la regla TOCA en octubre-1: el bloque
    /// completo del canal <c>O</c> del cuadro <c>Componente</c> (Vlr Servicio, Vlr Intereses, Total) y
    /// los tres agregados literales que se re-propagan. Son DATOS del caso §0.4, no coordenadas.
    /// </summary>
    private static readonly string[] EtiquetasReclasificadas =
    [
        EtiquetaDe("Componente", string.Empty, string.Empty, string.Empty, "Vlr Servicio", "O/OCCIDENTE"),
        EtiquetaDe("Componente", string.Empty, string.Empty, string.Empty, "Vlr Intereses", "O/OCCIDENTE"),
        EtiquetaDe("Componente", string.Empty, string.Empty, "O", "Total", "O/OCCIDENTE"),
        EtiquetaDe("Componente", string.Empty, "OCCIDENTE", "Total", string.Empty, null),
        EtiquetaDe("Componente", "Componente", "Total", string.Empty, string.Empty, null),
        EtiquetaDe(string.Empty, "Total", string.Empty, string.Empty, string.Empty, null),
    ];

    // ── I1 .. I4 (Plan 60, H3.1) ────────────────────────────────────────────────────────────────

    /// <summary>
    /// I1 (CA-R3): la salida de octubre-1 de ASE2 coincide con el oráculo administrativo <b>por
    /// separado</b> en <c>Deb/Cred</c> y <c>Deb/Cred-otros</c>, por ETIQUETA: canal <c>O</c>
    /// (0 / movido), <c>OCCIDENTE</c> (−6,15 / 6.759.348,81), <c>Componente</c> y gran <c>Total</c>
    /// (22,96 / 6.759.348,81). El comentario de divergencia declarada del Plan 59 queda retirado.
    /// </summary>
    [Fact]
    public void I1_2026101_ASE2_CanalO_Reclasificado()
    {
        var generacion = Generacion2026101.Value;
        var escrito = BloquePorEtiqueta(generacion.Salida, HojaR2, aseId: 2);
        var oraculo = BloquePorEtiqueta(Oraculo2026101, HojaR2, aseId: 2);

        var divergencias = new List<string>();
        foreach (var etiqueta in EtiquetasReclasificadas)
        {
            Assert.True(oraculo.ContainsKey(etiqueta), $"El oráculo no declara la etiqueta '{etiqueta}'.");
            Assert.True(escrito.ContainsKey(etiqueta), $"La salida no declara la etiqueta '{etiqueta}'.");

            var esperado = oraculo[etiqueta];
            var observado = escrito[etiqueta];
            if (Math.Abs(observado.Q - esperado.Q) > Tol)
            {
                divergencias.Add($"'{etiqueta}' · Deb/Cred: escrito={observado.Q} oráculo={esperado.Q} (Δ={observado.Q - esperado.Q}).");
            }

            if (Math.Abs(observado.R - esperado.R) > Tol)
            {
                divergencias.Add($"'{etiqueta}' · Deb/Cred-otros: escrito={observado.R} oráculo={esperado.R} (Δ={observado.R - esperado.R}).");
            }
        }

        Assert.True(divergencias.Count == 0, "El canal O de ASE2 no quedó reclasificado como el oráculo:\n" + string.Join("\n", divergencias));

        // Anclas explícitas del caso (mismos números que el oráculo, §0.4): el bloque del canal O
        // sale ENTERO de Deb/Cred y su importe aparece en Deb/Cred-otros.
        var servicio = EtiquetaDe("Componente", string.Empty, string.Empty, string.Empty, "Vlr Servicio", "O/OCCIDENTE");
        var intereses = EtiquetaDe("Componente", string.Empty, string.Empty, string.Empty, "Vlr Intereses", "O/OCCIDENTE");
        var cierreO = EtiquetaDe("Componente", string.Empty, string.Empty, "O", "Total", "O/OCCIDENTE");
        var occidente = EtiquetaDe("Componente", string.Empty, "OCCIDENTE", "Total", string.Empty, null);
        var componente = EtiquetaDe("Componente", "Componente", "Total", string.Empty, string.Empty, null);
        var granTotal = EtiquetaDe(string.Empty, "Total", string.Empty, string.Empty, string.Empty, null);

        Assert.Equal(0m, escrito[servicio].Q);
        Assert.Equal(54.00m, escrito[servicio].R);
        Assert.Equal(0m, escrito[intereses].Q);
        Assert.Equal(6_759_294.81m, escrito[intereses].R);
        Assert.Equal(0m, escrito[cierreO].Q);
        Assert.Equal(6_759_348.81m, escrito[cierreO].R);
        Assert.Equal(-6.15m, escrito[occidente].Q);
        Assert.Equal(6_759_348.81m, escrito[occidente].R);
        Assert.Equal(22.96m, escrito[componente].Q);
        Assert.Equal(22.96m, escrito[granTotal].Q);
    }

    /// <summary>
    /// I2 (CA-R3): <c>Deb/Cred-otros</c> queda en 0 en TODAS las demás filas del bloque ASE2, y su
    /// <c>Deb/Cred</c> sigue siendo el de la fuente (el oráculo lo confirma etiqueta por etiqueta).
    /// </summary>
    [Fact]
    public void I2_2026101_RestoDeFilas_R0()
    {
        var generacion = Generacion2026101.Value;
        var escrito = BloquePorEtiqueta(generacion.Salida, HojaR2, aseId: 2);
        var oraculo = BloquePorEtiqueta(Oraculo2026101, HojaR2, aseId: 2);

        var conOtros = escrito
            .Where(kv => Math.Abs(kv.Value.R) > Tol)
            .Select(kv => kv.Key)
            .OrderBy(k => k, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(
            EtiquetasReclasificadas.OrderBy(k => k, StringComparer.Ordinal).ToArray(),
            conOtros);

        var divergencias = new List<string>();
        foreach (var (etiqueta, observado) in escrito)
        {
            if (!oraculo.TryGetValue(etiqueta, out var esperado))
            {
                continue; // fila que la plantilla declara y el oráculo no trae (capacidad de plantilla)
            }

            if (Math.Abs(observado.Q - esperado.Q) > Tol)
            {
                divergencias.Add($"'{etiqueta}' · Deb/Cred: escrito={observado.Q} oráculo={esperado.Q}.");
            }

            if (Math.Abs(observado.R - esperado.R) > Tol)
            {
                divergencias.Add($"'{etiqueta}' · Deb/Cred-otros: escrito={observado.R} oráculo={esperado.R}.");
            }
        }

        Assert.True(divergencias.Count == 0, "Hay filas del bloque ASE2 fuera del caso reclasificado:\n" + string.Join("\n", divergencias));
    }

    /// <summary>
    /// I3 (CA-R3/D5): ASE1/3/4/5 quedan intactos — ningún canal suyo supera el umbral, así que
    /// <c>Deb/Cred-otros</c> queda en 0 en todo su bloque y <c>Deb/Cred</c> coincide con el oráculo.
    /// </summary>
    [Fact]
    public void I3_2026101_ASE1345_Intactos()
    {
        var generacion = Generacion2026101.Value;
        var divergencias = new List<string>();

        foreach (var aseId in new[] { 1, 3, 4, 5 })
        {
            var escrito = BloquePorEtiqueta(generacion.Salida, HojaR2, aseId);
            var oraculo = BloquePorEtiqueta(Oraculo2026101, HojaR2, aseId);

            foreach (var (etiqueta, observado) in escrito)
            {
                if (Math.Abs(observado.R) > Tol)
                {
                    divergencias.Add($"ASE {aseId} · '{etiqueta}' · Deb/Cred-otros={observado.R} (debía quedar en 0).");
                }

                if (oraculo.TryGetValue(etiqueta, out var esperado) && Math.Abs(observado.Q - esperado.Q) > Tol)
                {
                    divergencias.Add($"ASE {aseId} · '{etiqueta}' · Deb/Cred: escrito={observado.Q} oráculo={esperado.Q}.");
                }
            }
        }

        Assert.True(divergencias.Count == 0, "La reclasificación tocó un ASE que no debía:\n" + string.Join("\n", divergencias));
    }

    /// <summary>
    /// I4 (CA-R8): la traza del ASE2 de octubre-1 trae EXACTAMENTE los tres conceptos del canal <c>O</c>
    /// con su coordenada completa (hoja · fuente · ASE · sección · subsección/empresa · canal ·
    /// concepto) y los tres agregados reconciliados con su tramo; el grupo homónimo <c>O</c> del
    /// cuadro <c>Subs/Cont</c> (Q = 0) no produce ningún movimiento.
    /// </summary>
    [Fact]
    public void I4_2026101_Auditoria_TrazaExacta()
    {
        var generacion = Generacion2026101.Value;
        var fuente = Fuente(CarpetaAse(new[] { "Docs", "Prueba Octubre-1", "Insumos" }, 2), "RerpoteDetalleSaldosaFavor", 2);
        var crudo = new ExcelDataReaderWorkbookLeafInputReader().LeerDetalleR2(Insumos.Ase(2), fuente);

        var resultado = ReclasificadorDebCredR2.Aplicar(crudo);

        var conceptos = resultado.Movimientos
            .Where(m => m.Motivo.StartsWith(ReclasificadorDebCredR2.MotivoCierreCanal, StringComparison.Ordinal))
            .ToList();
        Assert.Equal(3, conceptos.Count);
        Assert.Equal(
            new[] { "Vlr Servicio", "Vlr Intereses", "Total" },
            conceptos.Select(m => m.Coordenada.Concepto).ToArray());
        Assert.Equal(
            new[] { 54.00m, 6_759_294.81m, 6_759_348.81m },
            conceptos.Select(m => m.MovidoQ).ToArray());

        foreach (var movimiento in conceptos)
        {
            var coordenada = movimiento.Coordenada;
            Assert.Equal(ReclasificadorDebCredR2.HojaRemAnticiposR2, coordenada.Hoja);
            Assert.Equal(fuente, coordenada.Fuente);
            Assert.Equal(2, coordenada.AseId);
            Assert.Equal("Componente", coordenada.Seccion);
            Assert.Equal("OCCIDENTE", coordenada.Subseccion);
            Assert.Equal("O", coordenada.Canal);
        }

        var agregados = resultado.Movimientos
            .Where(m => m.Motivo.StartsWith(ReclasificadorDebCredR2.MotivoAgregado, StringComparison.Ordinal))
            .ToList();
        Assert.Equal(3, agregados.Count);
        Assert.All(agregados, m => Assert.Equal(6_759_348.81m, m.MovidoQ));
        Assert.Contains(agregados, m => m.Coordenada.Subseccion == "OCCIDENTE" && m.Coordenada.Canal == "H,O,T");
        Assert.Contains(agregados, m => m.Coordenada.Concepto == "Componente");
        Assert.Contains(agregados, m => m.Coordenada.Seccion.Length == 0 && m.Coordenada.Concepto == "Total");

        // El grupo homónimo del 2.º cuadro (Q = 0) no genera NADA: la identidad es la instancia.
        Assert.DoesNotContain(resultado.Movimientos, m => m.Coordenada.Seccion == "Subs/Cont");
    }

    // ── G3 .. G6 (Plan 60, H3.2/H3.3) ───────────────────────────────────────────────────────────

    /// <summary>
    /// G3 (CA-R4): blast radius. El diff entre lo que declara la FUENTE y lo que quedó en la hoja es
    /// exactamente el par Q/R de los seis registros del caso; para toda fila tocada se conserva
    /// <c>Q + R</c> (es el invariante que mantiene cerrando <c>CONSOLIDADO</c> y <c>VALIDACION_TOTAL</c>).
    /// </summary>
    [Fact]
    public void G3_BlastRadius_2026101_SoloCeldasEsperadas()
    {
        var generacion = Generacion2026101.Value;
        var crudo = generacion.R2Ase2Crudo;
        var clasificado = ReclasificadorDebCredR2.Aplicar(crudo).Clasificado;

        // (1) El diff contra la fuente, a nivel de MODELO: solo las seis etiquetas del caso.
        var canales = CanalesEfectivos(crudo);
        var secciones = SeccionesPorFila(crudo);
        var tocadas = new List<string>();
        for (var i = 0; i < crudo.Filas.Count; i++)
        {
            var antes = crudo.Filas[i];
            var despues = clasificado.Filas[i];
            var qAntes = antes.Valor(ColumnaDebCred) ?? 0m;
            var rAntes = antes.Valor(ColumnaDebCredOtros) ?? 0m;
            var qDespues = despues.Valor(ColumnaDebCred) ?? 0m;
            var rDespues = despues.Valor(ColumnaDebCredOtros) ?? 0m;

            // Conservación del par en TODA fila (no sólo en las tocadas).
            Assert.True(
                Math.Abs((qAntes + rAntes) - (qDespues + rDespues)) <= Tol,
                $"fila #{i + 1} de la fuente: Q+R antes={qAntes + rAntes} después={qDespues + rDespues}.");

            if (Math.Abs(qAntes - qDespues) <= Tol && Math.Abs(rAntes - rDespues) <= Tol)
            {
                continue;
            }

            tocadas.Add(EtiquetaDe(secciones[i], antes.A, antes.B, antes.C, antes.D, canales[i]));
        }

        Assert.Equal(
            EtiquetasReclasificadas.OrderBy(k => k, StringComparer.Ordinal).ToArray(),
            tocadas.OrderBy(k => k, StringComparer.Ordinal).ToArray());

        // (2) La hoja escrita es EXACTAMENTE el modelo clasificado (por etiqueta).
        var escrito = BloquePorEtiqueta(generacion.Salida, HojaR2, aseId: 2);
        var porEtiquetaClasificado = new Dictionary<string, (decimal Q, decimal R)>(StringComparer.Ordinal);
        for (var i = 0; i < clasificado.Filas.Count; i++)
        {
            porEtiquetaClasificado[EtiquetaDe(secciones[i], clasificado.Filas[i].A, clasificado.Filas[i].B, clasificado.Filas[i].C, clasificado.Filas[i].D, canales[i])] =
                (clasificado.Filas[i].Valor(ColumnaDebCred) ?? 0m, clasificado.Filas[i].Valor(ColumnaDebCredOtros) ?? 0m);
        }

        foreach (var (etiqueta, valor) in porEtiquetaClasificado)
        {
            Assert.True(escrito.TryGetValue(etiqueta, out var hoja), $"La hoja no trae la fila '{etiqueta}' del modelo.");
            Assert.True(
                Math.Abs(hoja.Q - valor.Q) <= Tol && Math.Abs(hoja.R - valor.R) <= Tol,
                $"'{etiqueta}': hoja=({hoja.Q}/{hoja.R}) modelo clasificado=({valor.Q}/{valor.R}).");
        }
    }

    /// <summary>
    /// G4 (CA-R5): ninguna fórmula del cuadro OPORTUNO fue escrita — el conjunto de celdas con
    /// <c>&lt;f&gt;</c> de las columnas <c>Deb/Cred</c>/<c>Deb/Cred-otros</c> de la hoja, con su texto,
    /// es IDÉNTICO al de la plantilla vigente. Y las celdas del bloque de detalle siguen siendo valores.
    /// </summary>
    [Fact]
    public void G4_FormulasIntactas_OportunoSigueSiendoFormula()
    {
        var generacion = Generacion2026101.Value;

        var plantilla = FormulasDe(PlantillaVigente, HojaR2.Nombre);
        var generado = FormulasDe(generacion.Salida, HojaR2.Nombre);

        Assert.NotEmpty(plantilla);
        Assert.True(
            plantilla.SetEquals(generado),
            "El conjunto de fórmulas del cuadro OPORTUNO cambió."
            + Environment.NewLine
            + "Faltan en la salida: "
            + string.Join(" · ", plantilla.Except(generado))
            + Environment.NewLine
            + "Sobran en la salida: "
            + string.Join(" · ", generado.Except(plantilla)));

        // El bloque de detalle NO trae fórmulas en esas columnas: son celdas de valor.
        var bloque = BloquePorEtiqueta(generacion.Salida, HojaR2, aseId: 2);
        Assert.NotEmpty(bloque);
    }

    /// <summary>
    /// G5 (CA-R5/§0.7): julio-1, julio-2 y agosto-2 no disparan en NINGÚN ASE (medido en las fuentes:
    /// ningún cierre de canal supera el umbral) ⇒ la regla es no-op y Q/R no cambian (R = 0 explícito).
    /// </summary>
    [Fact]
    public void G5_RegresionHistorica_SinCambios()
    {
        var periodos = new (string Periodo, string[] Carpeta)[]
        {
            ("2026071", new[] { "Docs", "Prueba Julio-1", "Insumos" }),
            ("2026072", new[] { "Docs", "Prueba Julio-2", "Insumos" }),
            ("2026082", new[] { "Docs", "Prueba Agosto-2", "Insumos" }),
        };

        var lector = new ExcelDataReaderWorkbookLeafInputReader();
        var disparos = new List<string>();

        foreach (var (codigo, carpeta) in periodos)
        {
            var periodo = Periodo.Parse(codigo);
            var insumos = Ruta(carpeta);
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var r2 = Fuente(CarpetaAse(insumos, aseId), "RerpoteDetalleSaldosaFavor", aseId);
                var crudo = lector.LeerDetalleR2(Insumos.Ase(aseId), r2);
                var resultado = ReclasificadorDebCredR2.Aplicar(crudo);

                if (resultado.Movimientos.Count > 0)
                {
                    disparos.Add($"{periodo.CodigoCompleto} · ASE {aseId}: {resultado.Movimientos.Count} movimiento(s).");
                    continue;
                }

                for (var i = 0; i < crudo.Filas.Count; i++)
                {
                    var q = crudo.Filas[i].Valor(ColumnaDebCred) ?? 0m;
                    Assert.Equal(q, resultado.Clasificado.Filas[i].Valor(ColumnaDebCred) ?? 0m);
                    Assert.Equal(0m, resultado.Clasificado.Filas[i].Valor(ColumnaDebCredOtros) ?? 0m);
                }
            }
        }

        Assert.True(disparos.Count == 0, "La regla disparó en un período histórico (§0.7):\n" + string.Join("\n", disparos));
    }

    /// <summary>
    /// G6 (D3): <c>Reversion Pagos R4</c> no declara los puertos <c>Deb/Cred</c> / <c>Deb/Cred-otros</c>
    /// —su última columna de valor es <c>Otros</c>—, así que la regla no corre sobre R4 (el propio tipo
    /// de entrada de <see cref="ReclasificadorDebCredR2"/> lo impide: sólo acepta el detalle de R2).
    /// </summary>
    [Fact]
    public void G6_Alcance_R4SinDebCred_NoCorre()
    {
        var r4 = WorkbookLeafCellMapDetalleR2R4.ColumnasR4.Select(c => c.Encabezado ?? string.Empty).ToArray();
        Assert.DoesNotContain(r4, encabezado => encabezado.Contains("Deb", StringComparison.OrdinalIgnoreCase));

        var r2 = WorkbookLeafCellMapDetalleR2R4.ColumnasR2.Select(c => c.Encabezado ?? string.Empty).ToArray();
        Assert.Contains(ColumnaDebCred, r2);
        Assert.Contains(ColumnaDebCredOtros, r2);

        // Y la hoja de la plantilla vigente tampoco declara esas columnas en su bloque de R4.
        var columnasR4 = ColumnasDeValorDeLaPlantilla(PlantillaVigente, HojaR4, aseId: 1);
        Assert.DoesNotContain(ColumnaDebCred, columnasR4);
        Assert.DoesNotContain(ColumnaDebCredOtros, columnasR4);
    }

    // ── Cobertura R4 vigente (sin reclasificación: R4 no tiene esos puertos) ────────────────────

    /// <summary>El desglose de R4 de octubre-1 coincide con el oráculo administrativo componente por componente.</summary>
    [Fact]
    public void Q1_2026101_DetalleR4_COINCIDE_CON_EL_ORACULO()
    {
        var generacion = Generacion2026101.Value;
        var divergencias = new List<string>();

        foreach (var leaf in generacion.Leafs)
        {
            var aseId = leaf.Ase.Id;
            var escrito = GranTotalDelBloque(generacion.Salida, HojaR4, aseId);
            var oraculo = GranTotalDelBloque(Oraculo2026101, HojaR4, aseId);

            var comparados = 0;
            foreach (var (etiqueta, valor) in escrito)
            {
                if (!oraculo.TryGetValue(etiqueta, out var esperado))
                {
                    continue;
                }

                comparados++;
                if (Math.Abs(valor - esperado) > Tol)
                {
                    divergencias.Add($"{HojaR4.Nombre} ASE {aseId} · '{etiqueta}': escrito={valor} oráculo={esperado} (Δ={valor - esperado}).");
                }
            }

            Assert.True(comparados > 0, $"{HojaR4.Nombre} ASE {aseId}: no se comparó ningún componente contra el oráculo.");
        }

        Assert.True(divergencias.Count == 0, "El desglose-detalle de R4 no coincide con el oráculo administrativo:\n" + string.Join("\n", divergencias));
    }

    // ── Plan 59: Q1 de julio-1 escrito y cuadrado contra la fuente ─────────────────────────────

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

    // ── Generación única del período (reutilizada por I1..I4 y G3/G4) ──────────────────────────

    private sealed record Generacion(
        string Salida,
        IReadOnlyList<WorkbookLeafInputs> Leafs,
        DetalleR2AseInputs R2Ase2Crudo);

    private static readonly Lazy<Generacion> Generacion2026101 = new(CrearGeneracion2026101);

    private static Generacion CrearGeneracion2026101()
    {
        var periodo = Periodo.Parse("2026101");
        var insumos = Ruta("Docs", "Prueba Octubre-1", "Insumos");
        var (leafs, resultado) = ConstruirPeriodo(periodo, insumos);

        // Modelo CRUDO del ASE2 (observación fiel de la fuente) ANTES de escribir: el writer sustituye
        // el leaf por el modelo clasificado y el blast radius se mide contra este.
        var r2 = Fuente(CarpetaAse(insumos, 2), "RerpoteDetalleSaldosaFavor", 2);
        var crudo = new ExcelDataReaderWorkbookLeafInputReader().LeerDetalleR2(Insumos.Ase(2), r2);

        var directorio = NuevoDirectorio();
        var salida = Path.Combine(directorio, "salida.xlsx");
        new OpenXmlPlantillaWriter().GenerarWorkbook(PlantillaVigente, salida, resultado, leafs);

        return new Generacion(salida, leafs, crudo);
    }

    // ── Lectura por etiqueta (gramática de cierres + encabezados de la hoja) ────────────────────

    /// <summary>
    /// Bloque de la hoja indexado por ETIQUETA semántica: <c>sección | A | B | C | D | canal efectivo</c>
    /// (el canal efectivo sale de <see cref="CanalEfectivoDetalleR2R4"/>, la MISMA clave de colocación
    /// de producción, y desambigua las filas sin etiqueta propia dentro de su grupo). Nunca se usa el
    /// número de fila: el oráculo administrativo está sobre otra geometría de plantilla (§0.6).
    /// </summary>
    private static Dictionary<string, (decimal Q, decimal R)> BloquePorEtiqueta(
        string ruta,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        int aseId)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hoja, aseId);
        var filas = FilasDe(workbookPart, hoja);
        var columnas = ColumnasDeValor(workbookPart, filas, bloque)
            .ToDictionary(c => c.Etiqueta, c => c.Columna, StringComparer.OrdinalIgnoreCase);

        Assert.True(columnas.ContainsKey(ColumnaDebCred), $"La hoja '{hoja.Nombre}' no declara '{ColumnaDebCred}'.");
        Assert.True(columnas.ContainsKey(ColumnaDebCredOtros), $"La hoja '{hoja.Nombre}' no declara '{ColumnaDebCredOtros}'.");

        var rotulos = RotulosDelBloque(workbookPart, filas, bloque);
        var secciones = Secciones(rotulos);
        var canales = CanalesEfectivos(rotulos);

        var mapa = new Dictionary<string, (decimal Q, decimal R)>(StringComparer.Ordinal);
        for (var i = 0; i < rotulos.Count; i++)
        {
            var etiqueta = EtiquetaDe(secciones[i], rotulos[i], canales[i]);
            var fila = PrimeraFilaDeEtiquetas(bloque) + i;
            Assert.True(
                mapa.TryAdd(
                    etiqueta,
                    (ValorDe(workbookPart, filas, fila, columnas[ColumnaDebCred]),
                     ValorDe(workbookPart, filas, fila, columnas[ColumnaDebCredOtros]))),
                $"La etiqueta '{etiqueta}' no es única en el bloque del ASE {aseId} de '{hoja.Nombre}'.");
        }

        return mapa;
    }

    /// <summary>Conjunto (referencia, texto) de las celdas con fórmula de las columnas del par.</summary>
    private static HashSet<string> FormulasDe(string ruta, string nombreHoja)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, nombreHoja);
        var tabla = worksheet.Elements<SheetData>().FirstOrDefault()
            ?? throw new InvalidOperationException($"La hoja '{nombreHoja}' no tiene SheetData.");

        var celdas = new HashSet<string>(StringComparer.Ordinal);
        foreach (var fila in tabla.Elements<Row>())
        {
            foreach (var celda in fila.Elements<Cell>())
            {
                if (celda.CellFormula is null || celda.CellReference?.Value is not { } referencia)
                {
                    continue;
                }

                var columna = ColumnaDe(referencia);
                if (!string.Equals(columna, "Q", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(columna, "R", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                celdas.Add($"{referencia} = {celda.CellFormula.Text}");
            }
        }

        return celdas;
    }

    private static IReadOnlyList<CanalEfectivoDetalleR2R4.Rotulo> RotulosDelBloque(
        WorkbookPart workbookPart,
        IReadOnlyDictionary<int, Row> filas,
        DelimitadorBloquesHoja.BloqueHoja bloque)
    {
        var rotulos = new List<CanalEfectivoDetalleR2R4.Rotulo>();
        for (var fila = PrimeraFilaDeEtiquetas(bloque); fila <= bloque.GranTotal; fila++)
        {
            rotulos.Add(new CanalEfectivoDetalleR2R4.Rotulo(
                TextoDestino(workbookPart, filas, "A", fila),
                TextoDestino(workbookPart, filas, "B", fila),
                TextoDestino(workbookPart, filas, "C", fila),
                TextoDestino(workbookPart, filas, "D", fila)));
        }

        return rotulos;
    }

    /// <summary>
    /// Primera fila de DATOS del bloque: la siguiente a su fila de encabezados de columna. Sale de la
    /// estructura que la hoja declara (nunca de una ventana congelada).
    /// </summary>
    private static int PrimeraFilaDeEtiquetas(DelimitadorBloquesHoja.BloqueHoja bloque) =>
        bloque.EncabezadoColumnas + 1;

    /// <summary>
    /// Sección que encierra cada fila del bloque: la etiqueta <c>A</c> del primer cierre de sección
    /// (<c>A</c> no vacía con <c>B = Total</c>) en o por debajo de la fila. Es la MISMA derivación que
    /// usa la regla para su coordenada; el gran total no queda dentro de ninguna sección.
    /// </summary>
    private static IReadOnlyList<string> Secciones(IReadOnlyList<CanalEfectivoDetalleR2R4.Rotulo> rotulos)
    {
        var secciones = new string[rotulos.Count];
        for (var i = 0; i < rotulos.Count; i++)
        {
            secciones[i] = string.Empty;
            for (var j = i; j < rotulos.Count; j++)
            {
                if (rotulos[j].A.Trim().Length > 0 && EsTotal(rotulos[j].B))
                {
                    secciones[i] = rotulos[j].A.Trim();
                    break;
                }
            }
        }

        return secciones;
    }

    private static IReadOnlyList<string?> CanalesEfectivos(IReadOnlyList<CanalEfectivoDetalleR2R4.Rotulo> rotulos) =>
        Enumerable.Range(0, rotulos.Count)
            .Select(i => CanalEfectivoDetalleR2R4.CanalEfectivo(
                rotulos,
                i,
                WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR2,
                WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR2,
                WorkbookLeafCellMapDetalleR2R4.ColumnasSegundoCierreR2))
            .ToList();

    private static IReadOnlyList<string?> CanalesEfectivos(DetalleR2AseInputs detalle) =>
        CanalesEfectivos(detalle.Filas
            .Select(f => new CanalEfectivoDetalleR2R4.Rotulo(f.A, f.B, f.C, f.D))
            .ToList());

    private static IReadOnlyList<string> SeccionesPorFila(DetalleR2AseInputs detalle) =>
        Secciones(detalle.Filas.Select(f => new CanalEfectivoDetalleR2R4.Rotulo(f.A, f.B, f.C, f.D)).ToList());

    /// <summary>Etiqueta semántica de una fila: sección + firma + canal efectivo (nunca la fila).</summary>
    private static string EtiquetaDe(string seccion, CanalEfectivoDetalleR2R4.Rotulo rotulo, string? canal) =>
        EtiquetaDe(seccion, rotulo.A, rotulo.B, rotulo.C, rotulo.D, canal);

    private static string EtiquetaDe(string seccion, string a, string b, string c, string d, string? canal) =>
        $"{seccion}|{a.Trim()}|{b.Trim()}|{c.Trim()}|{d.Trim()}|{canal}";

    private static bool EsTotal(string texto) =>
        string.Equals(texto.Trim(), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase);

    private static string TextoDestino(
        WorkbookPart workbookPart,
        IReadOnlyDictionary<int, Row> filas,
        string columna,
        int fila)
    {
        if (!filas.TryGetValue(fila, out var row))
        {
            return string.Empty;
        }

        var referencia = columna + fila.ToString(CultureInfo.InvariantCulture);
        var celda = row.Elements<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        return celda is null ? string.Empty : TextoDe(workbookPart, celda).Trim();
    }

    private static IReadOnlyList<string> ColumnasDeValorDeLaPlantilla(
        string ruta,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        int aseId)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hoja, aseId);
        return ColumnasDeValor(workbookPart, FilasDe(workbookPart, hoja), bloque)
            .Select(c => c.Etiqueta)
            .ToList();
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
            var carpetaAse = CarpetaAse(carpetaPeriodo, i);
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

    private static string CarpetaAse(string carpetaPeriodo, int aseId) =>
        Directory.GetDirectories(carpetaPeriodo)
            .First(d => Path.GetFileName(d).StartsWith($"{aseId}-", StringComparison.Ordinal));

    private static string CarpetaAse(string[] partes, int aseId) => CarpetaAse(Ruta(partes), aseId);

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
