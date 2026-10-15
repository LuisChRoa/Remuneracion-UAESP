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
/// Plan 41 (T1, R-SEC-1/R-SEC-2/R-SEC-3/R-SEC-4 — D-A/D-H): resolución del espejo R1 por SECCIÓN.
///
/// La sección de una fila-dato de empresa la fija la frontera inmediatamente ARRIBA
/// (<see cref="SeccionEspejoR1.PorFronteraArriba"/>: <c>Mes</c>→Oportuno, <c>Aplicacion</c>→Aplicación,
/// sin frontera→1.ª). El escritor values-only por slot (<see cref="EscritorValoresSlotR1"/>) agrupa por
/// (firma + sección) y consume dentro de la sección en orden — cerrando el defecto E1-2 del Plan 40
/// (7 pares empresa×ASE que el cursor físico dejaba cortos/cero).
///
/// Oráculo: el valor del MANUAL del administrativo en SU ancla, por (empresa, ASE) — NUNCA se comparan
/// direcciones legacy↔oficial (R-TRAMPA-ANCLA). El oficial es el input; el <c>&lt;f&gt;</c> del oficial
/// queda intacto (solo se escriben <c>&lt;v&gt;</c> en los slots). OCCIDENTE ASE5 es una divergencia de
/// ESTRUCTURA (mini-T0): queda DECLARADO, nunca verde.
/// </summary>
public sealed class SeccionEspejoR1Tests
{
    private const string HojaR1 = "Reporte Componentes R1";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental",
        "2-Lime",
        "3-Ciudad Limpia",
        "4-Bogota Limpia",
        "5-Área Limpia",
    ];

    // Ancla oficial por (empresa, ASE) — REMUNERACION_*!D9:D13 -> 'Reporte Componentes R1'!F{celda}.
    private static readonly Dictionary<string, string[]> AnclasPorEmpresa = new(StringComparer.Ordinal)
    {
        ["REMUNERACION Reciprocidad EAAB"] = ["F80", "F259", "F420", "F573", "F685"],
        ["REMUNERACION_ENEL"] = ["F85", "F264", "F425", "F578", "F690"],
        ["REMUNERACION_ENERBIT"] = ["F90", "F269", "F430", "F583", "F695"],
        ["REMUNERACION_OCCIDENTE_ Directa"] = ["F95", "F274", "F435", "F588", "F700"],
        ["REMUNERACION_EAAB-CL"] = ["F100", "F279", "F440", "F593", "F705"],
    };

    private const decimal Tolerancia = 0.5m;

    // ── Unitarios del predicado (sin I/O) ────────────────────────────────────────────────────────

    private static FilaEspejoR1 F(string a = "", string b = "", string c = "", string d = "", string e = "") =>
        new() { A = a, B = b, C = c, D = d, E = e };

    [Fact]
    public void Seccion_FronterasArriba_TipoYOrdinal()
    {
        var filas = new List<FilaEspejoR1>
        {
            F(c: "ENEL", d: "Total"),                          // 0 1.ª sección
            F(b: "Aplicacion nuevos x reversion", c: "Total"), // 1 FRONTERA Aplicacion #1
            F(c: "OCCIDENTE", d: "Total"),                     // 2 Aplicación #1
            F(b: "Mes", c: "Total"),                           // 3 FRONTERA Mes #1
            F(c: "ENEL", d: "Total"),                          // 4 Oportuno #1
            F(b: "Aplicacion nuevos x reversion", c: "Total"), // 5 FRONTERA Aplicacion #2
            F(c: "ENEL", d: "Total"),                          // 6 Aplicación #2
            F(b: "Mes", c: "Total"),                           // 7 FRONTERA Mes #2
            F(c: "OCCIDENTE", d: "Total"),                     // 8 Oportuno #2
        };

        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Primera, 0), SeccionEspejoR1.PorFronteraArriba(filas, 0));
        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 1), SeccionEspejoR1.PorFronteraArriba(filas, 2));
        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1), SeccionEspejoR1.PorFronteraArriba(filas, 4));
        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 2), SeccionEspejoR1.PorFronteraArriba(filas, 6));
        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2), SeccionEspejoR1.PorFronteraArriba(filas, 8));
    }

    // ── S1/S2: los 7 pares empresa×ASE al valor del manual (±0.5) ────────────────────────────────

    [Theory]
    // Reciprocidad ASE2 (F259 = F200+F160-L160) y ASE4 (F573 = F540+F500-L500).
    [InlineData("REMUNERACION Reciprocidad EAAB", 2, 11411240.00)]
    [InlineData("REMUNERACION Reciprocidad EAAB", 4, 5616850.00)]
    // ENEL ASE3 (F425) y ASE5 (F690).
    [InlineData("REMUNERACION_ENEL", 3, 13997194578.87)]
    [InlineData("REMUNERACION_ENEL", 5, 11701157844.90)]
    // OCCIDENTE ASE1 (F95) y ASE3 (F435).
    [InlineData("REMUNERACION_OCCIDENTE_ Directa", 1, 779194409.36)]
    [InlineData("REMUNERACION_OCCIDENTE_ Directa", 3, 1229014628.26)]
    public void T1_CierraParEmpresaAse_AlValorDelManual(string hojaEmpresa, int aseId, double esperadoManual)
    {
        var valor = EscribirEspejoYEvaluarAncla(aseId, hojaEmpresa);
        Assert.InRange(valor - (decimal)esperadoManual, -Tolerancia, Tolerancia);
    }

    // ── R-SEC-4 / S7: los pares que hoy cierran por casualidad quedan blindados ──────────────────

    [Theory]
    [InlineData("REMUNERACION_ENEL", 1, 16560599150.50)]
    [InlineData("REMUNERACION_ENEL", 2, 19023525223.11)]
    [InlineData("REMUNERACION_ENEL", 4, 6584787380.47)]
    [InlineData("REMUNERACION_OCCIDENTE_ Directa", 2, 1317332077.61)]
    [InlineData("REMUNERACION_OCCIDENTE_ Directa", 4, 457088264.26)]
    [InlineData("REMUNERACION_ENERBIT", 5, 26864220.00)]
    public void T1_IdentidadParesQueCierran_AlValorDelManual(string hojaEmpresa, int aseId, double esperadoManual)
    {
        var valor = EscribirEspejoYEvaluarAncla(aseId, hojaEmpresa);
        Assert.InRange(valor - (decimal)esperadoManual, -Tolerancia, Tolerancia);
    }

    // ── OCCIDENTE ASE5: divergencia de ESTRUCTURA (mini-T0) — DECLARADO, nunca verde ─────────────

    [Fact]
    public void T1_OccidenteAse5_DivergenciaEstructural_Declarada()
    {
        // El oficial F700 = F647+F673+F612-L612-L647 (3 términos) NO corresponde al manual
        // F578 = F530+F551-L530 (2 términos). Colocando por sección (src13→612, src32→647,
        // src53→673) el oficial da 305.407.140,43 mientras el manual da 304.989.620,81: la
        // diferencia es EXACTAMENTE la fila de la 1.ª sección de OCCIDENTE (417.519,62). Es un
        // hallazgo de plantilla para la UAESP (no se resuelve por supuesto, D-E).
        var valorOficialPorSeccion = EscribirEspejoYEvaluarAncla(5, "REMUNERACION_OCCIDENTE_ Directa");

        Assert.InRange(valorOficialPorSeccion - 305407140.43m, -Tolerancia, Tolerancia);

        // El valor del manual difiere (no se cierra por colocación): divergencia declarada.
        Assert.True(Math.Abs(valorOficialPorSeccion - 304989620.81m) > Tolerancia,
            "OCCIDENTE ASE5: el oficial-por-sección NO debe coincidir con el manual (divergencia estructural declarada).");
    }

    // ── S3 / R-SEC-2: fila sin grupo de sección → fail-fast que nombra ASE+empresa+fila ──────────

    [Fact]
    public void T1_FilaSinSlotLibre_FailFastSinRemap()
    {
        // Una fila-dato de empresa SIN slot libre en su firma (más filas ENEL que slots ENEL del
        // ASE1: el oficial declara 5) → ERR-SIN-SLOT que nombra ASE + empresa + sección; cero
        // escritura parcial y NUNCA remap a otra empresa.
        using var salida = new CopiaTemporal();
        var filas = new List<FilaEspejoR1>();
        for (var k = 0; k < 6; k++)
        {
            filas.Add(ConTotal("ENEL"));
        }

        var bloque = new BloqueEspejoAseInputs
        {
            Ase = AseFactory.DesdeId(1),
            TieneColumnaEspeciales = false,
            Filas = filas,
        };

        using var documento = SpreadsheetDocument.Open(salida.Ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var ex = Assert.Throws<CalculoInvalidoException>(() => EscritorValoresSlotR1.Escribir(parte, bloque));

        Assert.Equal(CodigoError.SinSlot, ex.Codigo);
        Assert.Contains("ASE 1", ex.Message, StringComparison.Ordinal);
        Assert.Contains("ENEL", ex.Message, StringComparison.Ordinal);

        static FilaEspejoR1 ConTotal(string empresa) => new()
        {
            C = empresa,
            D = "Total",
            ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase) { ["Total"] = 5m },
        };
    }

    // ── E1-2: fila con rótulo SINÓNIMO resuelve por el CATÁLOGO (no por posición) ────────────────

    [Fact]
    public void T1_FilaSinonima_ResuelvePorCatalogoEnSuSeccion()
    {
        // La fuente trae el rótulo vigente 'RECIPROCIDAD' y el oficial declara el slot 'NUEVO ESQUEMA'
        // (misma EFC; SinonimosEmpresaR1, Plan 34). La fila, en su sección (frontera Mes ARRIBA →
        // Oportuno #1), resuelve por el CATÁLOGO contra el slot NUEVO ESQUEMA del ASE2 (F135) — nunca
        // por posición ni por un fallback. Prueba que la resolución de sección consume el catálogo.
        using var salida = new CopiaTemporal();
        var bloque = new BloqueEspejoAseInputs
        {
            Ase = AseFactory.DesdeId(2),
            TieneColumnaEspeciales = false,
            Filas =
            [
                new FilaEspejoR1 { B = "Mes", C = "Total" }, // frontera Mes #1 → sección Oportuno #1
                ConTotal("RECIPROCIDAD", 9876543.21m),
            ],
        };

        using (var documento = SpreadsheetDocument.Open(salida.Ruta, true))
        {
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var resultado = EscritorValoresSlotR1.Escribir(parte, bloque);
            parte.Workbook?.Save();
            Assert.Equal(2, resultado.FilasColocadas);
        }

        // F135 = slot NUEVO ESQUEMA (Oportuno #1) del ASE2 en la geometría oficial.
        Assert.Equal(9876543.21m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F135"));

        static FilaEspejoR1 ConTotal(string empresa, decimal total) => new()
        {
            C = empresa,
            D = "Total",
            ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase) { ["Total"] = total },
        };
    }

    // ── E1-3: fila sin sección resoluble → ERR-SIN-SLOT (nombra fila + rótulo + sección) ──────────

    [Fact]
    public void T1_FilaSinSeccionResoluble_ErrSinSlotConRotuloYSeccion()
    {
        // Una fila-dato de empresa cuya (firma, SECCIÓN) no tiene slot en el contrato T0 — aquí
        // 'NUEVO ESQUEMA' en la 1.ª sección del ASE2: el oficial NO declara ese slot (F110:F119 trae
        // solo ENEL F112 y OCCIDENTE F119) → ERR-SIN-SLOT que nombra ASE + rótulo de empresa +
        // sección esperada. Es información accionable, no un fallo mudo; cero colocación por posición.
        using var salida = new CopiaTemporal();
        var bloque = new BloqueEspejoAseInputs
        {
            Ase = AseFactory.DesdeId(2),
            TieneColumnaEspeciales = false,
            Filas = [ConTotal("NUEVO ESQUEMA", 5m)],
        };

        using var documento = SpreadsheetDocument.Open(salida.Ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var ex = Assert.Throws<CalculoInvalidoException>(() => EscritorValoresSlotR1.Escribir(parte, bloque));

        Assert.Equal(CodigoError.SinSlot, ex.Codigo);
        Assert.Contains("ERR-SIN-SLOT", ex.Message, StringComparison.Ordinal);
        Assert.Contains("ASE 2", ex.Message, StringComparison.Ordinal);
        Assert.Contains("NUEVO ESQUEMA", ex.Message, StringComparison.Ordinal);
        Assert.Contains("1.ª sección", ex.Message, StringComparison.Ordinal);

        static FilaEspejoR1 ConTotal(string empresa, decimal total) => new()
        {
            C = empresa,
            D = "Total",
            ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase) { ["Total"] = total },
        };
    }

    // ── E1-4: evidencia del hallazgo al quitar el fallback (fuente Q1 de ASE2, rótulo NUEVO ESQUEMA) ─

    [Fact]
    public void T1_Ase2Q1_NuevoEsquemaEnPrimeraSeccion_ErrSinSlotSinFallback()
    {
        // Evidencia E1-4 (Plan 41, Enmienda E-1): la fuente Q1 de ASE2 (Lime) trae una fila
        // 'NUEVO ESQUEMA' en su 1.ª sección; el oficial NO declara un slot NUEVO ESQUEMA (ni
        // RECIPROCIDAD) en su 1.ª sección (F110:F119 solo ENEL F112 / OCCIDENTE F119) y el catálogo
        // NO puede crear el slot (solo iguala rótulos) ⇒ ERR-SIN-SLOT accionable. NO se reintroduce
        // el fallback por cursor (E1-1). La causa se reporta al Ingeniero/UAESP.
        var bloque = new ExcelDataReaderWorkbookLeafInputReader().LeerEspejoR1(AseFactory.DesdeId(2), FuenteR1JulioQ1(2));
        using var salida = new CopiaTemporal();

        using var documento = SpreadsheetDocument.Open(salida.Ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var ex = Assert.Throws<CalculoInvalidoException>(() => EscritorValoresSlotR1.Escribir(parte, bloque));

        Assert.Equal(CodigoError.SinSlot, ex.Codigo);
        Assert.Contains("ASE 2", ex.Message, StringComparison.Ordinal);
        Assert.Contains("NUEVO ESQUEMA", ex.Message, StringComparison.Ordinal);
        Assert.Contains("1.ª sección", ex.Message, StringComparison.Ordinal);
    }

    // ── Helpers ─────────────────────────────────────────────────────────────────────────────────

    private static decimal EscribirEspejoYEvaluarAncla(int aseId, string hojaEmpresa)
    {
        var bloque = new ExcelDataReaderWorkbookLeafInputReader().LeerEspejoR1(AseFactory.DesdeId(aseId), FuenteR1Julio(aseId));

        using var salida = new CopiaTemporal();
        using (var documento = SpreadsheetDocument.Open(salida.Ruta, true))
        {
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            EscritorValoresSlotR1.Escribir(parte, bloque);
            parte.Workbook?.Save();
        }

        var ancla = AnclasPorEmpresa[hojaEmpresa][aseId - 1];
        return EvaluarAncla(salida.Ruta, ancla);
    }

    /// <summary>
    /// Evalúa el ancla <paramref name="celdaAncla"/> del R1 desde los VALORES colocados: si la celda
    /// trae fórmula (suma/resta de refs F/L) la resuelve; si es valor plano lo devuelve. El oráculo
    /// son los valores escritos, nunca el caché de una fórmula no colocada.
    /// </summary>
    private static decimal EvaluarAncla(string ruta, string celdaAncla)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = Hoja(parte, HojaR1);
        var valores = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        string? formula = null;
        foreach (var cell in hoja.Descendants<Cell>())
        {
            var referencia = cell.CellReference?.Value ?? string.Empty;
            if (cell.CellFormula is not null)
            {
                if (string.Equals(referencia, celdaAncla, StringComparison.OrdinalIgnoreCase))
                {
                    formula = cell.CellFormula.Text;
                }

                continue;
            }

            if (cell.CellValue is not null && decimal.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var v))
            {
                valores[referencia] = v;
            }
        }

        if (formula is null)
        {
            return valores.GetValueOrDefault(celdaAncla, 0m);
        }

        return EvaluarSuma(formula, valores);
    }

    private static decimal EvaluarSuma(string formula, IReadOnlyDictionary<string, decimal> valores)
    {
        // Formas de ancla: suma/resta de refs F/L (p. ej. F663+F633-L633; con la rareza F112--L112).
        var normalizada = formula.Replace("--", "+", StringComparison.Ordinal);
        var total = 0m;
        var signo = 1;
        var i = 0;
        while (i < normalizada.Length)
        {
            var ch = normalizada[i];
            if (ch == '+')
            {
                signo = 1;
                i++;
            }
            else if (ch == '-')
            {
                signo = -1;
                i++;
            }
            else
            {
                var inicio = i;
                while (i < normalizada.Length && (char.IsAsciiLetter(normalizada[i]) || char.IsAsciiDigit(normalizada[i])))
                {
                    i++;
                }

                var celda = normalizada[inicio..i];
                total += signo * valores.GetValueOrDefault(celda, 0m);
                signo = 1;
            }
        }

        return total;
    }

    private static Worksheet Hoja(WorkbookPart parte, string nombre)
    {
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;
    }

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (plantilla estandarizada única).");
    }

    private static string Plantilla() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string FuenteR1Julio(int aseId)
    {
        var carpeta = Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Insumos", CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta explícita: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("Recaudoporcomponente", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Recaudoporcomponente del ASE{aseId} en {carpeta}");
        return archivo!;
    }

    /// <summary>Fuente R1 del período Q1 (<c>Docs/Prueba Julio-1/Insumos</c>) por ruta explícita.</summary>
    private static string FuenteR1JulioQ1(int aseId)
    {
        var carpeta = Path.Combine(Raiz(), "Docs", "Prueba Julio-1", "Insumos", CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta explícita: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("Recaudoporcomponente", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Recaudoporcomponente del ASE{aseId} en {carpeta}");
        return archivo!;
    }

    private sealed class CopiaTemporal : IDisposable
    {
        private readonly string _dir;

        public CopiaTemporal()
        {
            _dir = Path.Combine(Path.GetTempPath(), "p41-t1-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
            Ruta = Path.Combine(_dir, "salida.xlsx");
            File.Copy(Plantilla(), Ruta, overwrite: false);
        }

        public string Ruta { get; }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_dir))
                {
                    Directory.Delete(_dir, recursive: true);
                }
            }
            catch
            {
                // best-effort
            }
        }
    }
}
