using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 31 (T1, R-G-1 / S1): gate de paridad <b>workbook-vs-dominio</b> de los totales visibles del
/// bloque R1 de la hoja <c>Reporte Componentes R1</c>.
///
/// Por ASE localiza, por FIRMA de etiqueta (nunca por dirección congelada), los tres visibles del
/// contrato del plan:
///   - TOT_OPT: fila <c>C='TOTAL' ∧ D='OPORTUNO'</c> (columna F).
///   - total TDF: misma fila, columna G.
///   - EXTEMP: fila <c>D='EXTEMPORANEO'</c> posterior a la TOT_OPT (columna F).
/// Lee el TEXTO de fórmula (<c>&lt;f&gt;</c>) y los literales del propio workbook (BCL puro ZIP+XML;
/// PROHIBIDO usar el caché <c>&lt;v&gt;</c> de una celda con fórmula: son stale por diseño, método
/// D-E del Plan 29) y evalúa la suma lineal re-derivable. Compara ±0.5 contra los agregados de
/// dominio por firma (<see cref="WorkbookLeafInputsR1.TotalOportunoEsperadoPorAse"/> y
/// <see cref="WorkbookLeafInputsR1.ExtemporaneoEsperadoPorAse"/>).
///
/// El total TDF G no tiene agregado de dominio: se valida ESTRUCTURALMENTE contra la firma
/// <c>Mes/Total</c> del bloque. El manual (E2 congelado) compone <c>TDF = ΣG(Mes) − G(última Mes)</c>
/// por orden físico (excluye la fila Mes de mayor número de fila); el gate exige exactamente ese
/// conjunto de referencias. Nota de T1: el plan §2.3 suponía <c>ΣG(todas las Mes)</c>, pero lo que
/// dice el disco (los dos manuales del administrativo) MANDA: se fija la forma real.
///
/// Fail-fast (archivo + hoja + celda + necesidad) si una referencia de la fórmula no resuelve a un
/// literal del workbook, si apunta a otra fórmula, o si falta el bloque/visible.
/// </summary>
public sealed class ValidadorTotalesR1Workbook
{
    /// <summary>Hoja objeto del gate (misma constante que el espejo R1).</summary>
    public const string HojaR1 = "Reporte Componentes R1";

    /// <summary>Tolerancia de cierre del proyecto (±0.5).</summary>
    public const decimal Tolerancia = 0.5m;

    private static readonly Regex ReferenciaCelda = new(
        @"^(?<col>[A-Za-z]{1,3})(?<fila>\d+)$",
        RegexOptions.Compiled);

    /// <summary>
    /// Término lineal de una fórmula de total R1: signo opcional + referencia A1 (con <c>$</c>
    /// opcionales). Las fórmulas visibles R1 son sumas lineales puras (0 funciones, 0 paréntesis).
    /// </summary>
    private static readonly Regex TerminoFormula = new(
        @"(?<signo>[+-]?)(?<ref>\$?[A-Za-z]{1,3}\$?\d+)",
        RegexOptions.Compiled);

    /// <summary>
    /// Valida los visibles R1 del workbook generado contra el dominio por firma. Devuelve las
    /// divergencias (vacío = PASS). Fail-fast si el workbook/hoja/bloque/visible no se puede leer.
    /// </summary>
    public ResultadoValidacionTotalesR1 Validar(string rutaWorkbook, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        ArgumentNullException.ThrowIfNull(leafs);

        var celdas = LeerCeldas(rutaWorkbook);
        var filas = FilasOrdenadas(celdas);
        var divergencias = new List<DivergenciaTotalesR1>();

        foreach (var leaf in leafs.OrderBy(l => l.Ase.Id))
        {
            ValidarAse(celdas, filas, leaf, divergencias);
        }

        return new ResultadoValidacionTotalesR1(divergencias);
    }

    /// <summary>
    /// Igual que <see cref="Validar"/> pero fail-fast si hay cualquier divergencia (uso como gate
    /// del pipeline). El mensaje nombra hoja + celda + esperado + evaluado.
    /// </summary>
    public void Exigir(string rutaWorkbook, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        var resultado = Validar(rutaWorkbook, leafs);
        if (!resultado.EsValido)
        {
            throw new CalculoInvalidoException(CodigoError.Validacion, resultado.Mensaje);
        }
    }

    /// <summary>
    /// Plan 31 (T2, R-B-3/D-B): sucesor del gate de fórmulas protegidas que el writer salta con el
    /// espejo desplazado (Δ≠0). Valida la <b>forma recompuesta</b> (texto-fórmula por firma) de los
    /// visibles R1 contra las filas REALES del propio workbook: TOT_OPT F = refs F(Mes) ∪ L(Mes
    /// menos la última); TDF G = refs G(Mes menos la última); EXTEMP F = refs F(Aplic) ∪ L(primera
    /// Aplic), o literal (sin Aplic). NO usa los fragmentos congelados a julio; NO lee así jamás un
    /// literal del dominio (es puramente estructural, complementario al gate numérico
    /// <see cref="Validar"/>). Fail-fast nombrando ASE + celda + conjunto real vs esperado.
    /// </summary>
    public void ExigirFormaRecompuesta(string rutaWorkbook, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        ArgumentNullException.ThrowIfNull(leafs);

        var celdas = LeerCeldas(rutaWorkbook);
        var filas = FilasOrdenadas(celdas);
        var errores = new List<string>();

        foreach (var leaf in leafs.OrderBy(l => l.Ase.Id))
        {
            ValidarFormaAse(celdas, filas, leaf.Ase, errores);
        }

        if (errores.Count > 0)
        {
            throw Fail($"la forma recompuesta de los visibles R1 no cierra:\n{string.Join("\n", errores)}");
        }
    }

    /// <summary>
    /// Plan 32 (T1, pieza c — R-G-1 / S1): gate-v2 ESTRUCTURAL del INTERIOR de <c>Reporte
    /// Componentes R1</c> (sub-visibles por empresa).
    ///
    /// Verifica, por CÓDIGO (texto de <c>&lt;f&gt;</c> y literales del propio workbook; PROHIBIDO usar
    /// el caché <c>&lt;v&gt;</c>), que cada celda interior congelada en <paramref name="esperados"/>
    /// coincida con el texto del manual del administrativo:
    ///   - <c>Formula</c> no nulo: la celda de la salida debe traer EXACTAMENTE ese texto de fórmula.
    ///   - <c>Formula</c> nulo (clase L-1: el manual no trae <c>&lt;f&gt;</c>): la salida NO debe ser
    ///     fórmula (literal/vacío), y si el fixture trae <c>ValorLiteral</c>, su literal debe coincidir.
    ///
    /// NO evalúa valores agregados (sin dominio nuevo, D-F): es complementario al comparador f-vs-f (la
    /// red numérica) y al gate numérico <see cref="Validar"/>. Fail-fast nombrando ASE + hoja + celda +
    /// esperado + real (con el conjunto de refs para lectura estructural).
    ///
    /// T1 lo deja en ROJO sobre la salida de agosto (el interior conserva el anclaje julio) y en VERDE
    /// sobre julio (identidad, D-E); T2 lo cierra al recomponer el interior por firma.
    /// </summary>
    public void ExigirFormaRecompuestaInterior(string rutaWorkbook, IReadOnlyList<InteriorR1Esperado> esperados)
    {
        ArgumentNullException.ThrowIfNull(esperados);

        var celdas = LeerCeldas(rutaWorkbook);
        var errores = new List<string>();

        foreach (var esperado in esperados.OrderBy(e => e.AseId).ThenBy(e => e.Celda, StringComparer.Ordinal))
        {
            ValidarCeldaInterior(celdas, esperado, errores);
        }

        if (errores.Count > 0)
        {
            throw Fail($"la forma interior de R1 no cierra:\n{string.Join("\n", errores)}");
        }
    }

    private static void ValidarCeldaInterior(
        IReadOnlyDictionary<string, CeldaR1Xlsx> celdas,
        InteriorR1Esperado esperado,
        List<string> errores)
    {
        var celda = celdas.GetValueOrDefault(esperado.Celda);

        if (esperado.Formula is null)
        {
            // Clase L-1: el manual no trae <f>; la salida no debe ser fórmula (y su literal debe
            // coincidir con el fixture cuando este lo declare).
            if (celda?.Formula is not null)
            {
                errores.Add($"ASE {esperado.AseId} {HojaR1}!{esperado.Celda} [{esperado.Etiqueta}]: esperado LITERAL (sin <f>, el manual no la trae), real fórmula [{celda.Formula}].");
                return;
            }

            if (esperado.ValorLiteral is { } literal && (celda?.Numero ?? 0m) != literal)
            {
                errores.Add($"ASE {esperado.AseId} {HojaR1}!{esperado.Celda} [{esperado.Etiqueta}]: esperado literal {Formatear(literal)}, real {(celda?.Numero is { } n ? Formatear(n) : "vacío")}.");
            }

            return;
        }

        if (celda?.Formula is null)
        {
            errores.Add($"ASE {esperado.AseId} {HojaR1}!{esperado.Celda} [{esperado.Etiqueta}]: esperado [{esperado.Formula}], real sin fórmula.");
            return;
        }

        if (!string.Equals(celda.Formula, esperado.Formula, StringComparison.Ordinal))
        {
            errores.Add($"ASE {esperado.AseId} {HojaR1}!{esperado.Celda} [{esperado.Etiqueta}]: esperado [{esperado.Formula}] (refs [{Refs(esperado.Formula)}]), real [{celda.Formula}] (refs [{Refs(celda.Formula)}]).");
        }
    }

    private static string Refs(string formula) =>
        string.Join(",",
            TerminoFormula.Matches(formula)
                .Select(m => NormalizarRef(m.Groups["ref"].Value))
                .OrderBy(r => r, StringComparer.OrdinalIgnoreCase));

    private static void ValidarFormaAse(
        IReadOnlyDictionary<string, CeldaR1Xlsx> celdas,
        IReadOnlyList<int> filas,
        Ase ase,
        List<string> errores)
    {
        var aseId = ase.Id;
        var nombre = Normalizar(ase.NombreCorto);

        var nameRow = filas.FirstOrDefault(r => string.Equals(Normalizar(Texto(celdas, "B" + r)), nombre, StringComparison.Ordinal));
        if (nameRow == 0)
        {
            errores.Add($"ASE {aseId}: no se halló el bloque (B='{ase.NombreCorto}').");
            return;
        }

        var totOptRow = filas
            .Where(r => r > nameRow)
            .FirstOrDefault(r => Normalizar(Texto(celdas, "C" + r)) == "TOTAL" && Normalizar(Texto(celdas, "D" + r)) == "OPORTUNO");
        if (totOptRow == 0)
        {
            errores.Add($"ASE {aseId}: no se halló el visible TOT_OPT.");
            return;
        }

        var extempRow = filas
            .Where(r => r > totOptRow)
            .FirstOrDefault(r => Normalizar(Texto(celdas, "D" + r)) == "EXTEMPORANEO");
        if (extempRow == 0)
        {
            errores.Add($"ASE {aseId}: no se halló el visible EXTEMP.");
            return;
        }

        var mesRows = filas
            .Where(r => r > nameRow && r < totOptRow
                && Normalizar(Texto(celdas, "B" + r)) == "MES"
                && Normalizar(Texto(celdas, "C" + r)) == "TOTAL")
            .OrderBy(r => r)
            .ToList();
        if (mesRows.Count == 0)
        {
            errores.Add($"ASE {aseId}: no se hallaron filas Mes/Total.");
            return;
        }

        var aplicRows = filas
            .Where(r => r > nameRow && r < totOptRow
                && Normalizar(Texto(celdas, "B" + r)).Contains("APLICACION NUEVOS X REVERSION", StringComparison.Ordinal)
                && Normalizar(Texto(celdas, "C" + r)) == "TOTAL")
            .OrderBy(r => r)
            .ToList();

        var esperadasTotOpt = new HashSet<string>(mesRows.Select(r => "F" + r), StringComparer.OrdinalIgnoreCase);
        if (mesRows.Count > 1)
        {
            foreach (var r in mesRows.Take(mesRows.Count - 1))
            {
                esperadasTotOpt.Add("L" + r);
            }
        }

        VerificarConjuntoRefs(celdas, "F" + totOptRow, esperadasTotOpt, aseId, "TOT_OPT", errores);

        var esperadasTdf = new HashSet<string>(
            (mesRows.Count > 1 ? mesRows.Take(mesRows.Count - 1) : mesRows).Select(r => "G" + r),
            StringComparer.OrdinalIgnoreCase);
        VerificarConjuntoRefs(celdas, "G" + totOptRow, esperadasTdf, aseId, "TDF", errores);

        var extempRef = "F" + extempRow;
        if (aplicRows.Count == 0)
        {
            if (celdas.GetValueOrDefault(extempRef)?.Formula is not null)
            {
                errores.Add($"ASE {aseId} {HojaR1}!{extempRef} [EXTEMP]: sin filas Aplicacion pero la celda es fórmula (debe ser literal 0).");
            }
        }
        else
        {
            var esperadasExtemp = new HashSet<string>(aplicRows.Select(r => "F" + r), StringComparer.OrdinalIgnoreCase) { "L" + aplicRows[0] };
            VerificarConjuntoRefs(celdas, extempRef, esperadasExtemp, aseId, "EXTEMP", errores);
        }
    }

    private static void VerificarConjuntoRefs(
        IReadOnlyDictionary<string, CeldaR1Xlsx> celdas,
        string celdaRef,
        HashSet<string> esperadas,
        int aseId,
        string etiqueta,
        List<string> errores)
    {
        var celda = celdas.GetValueOrDefault(celdaRef);
        if (celda?.Formula is null)
        {
            errores.Add($"ASE {aseId} {HojaR1}!{celdaRef} [{etiqueta}]: no es fórmula.");
            return;
        }

        var reales = new HashSet<string>(
            TerminoFormula.Matches(celda.Formula).Select(m => NormalizarRef(m.Groups["ref"].Value)),
            StringComparer.OrdinalIgnoreCase);
        if (!reales.SetEquals(esperadas))
        {
            errores.Add($"ASE {aseId} {HojaR1}!{celdaRef} [{etiqueta}]: refs [{string.Join(",", reales.OrderBy(r => r, StringComparer.OrdinalIgnoreCase))}] != esperadas [{string.Join(",", esperadas.OrderBy(r => r, StringComparer.OrdinalIgnoreCase))}].");
        }
    }

    // ── Validación por ASE ────────────────────────────────────────────────────────────────────

    private static void ValidarAse(
        IReadOnlyDictionary<string, CeldaR1Xlsx> celdas,
        IReadOnlyList<int> filas,
        WorkbookLeafInputs leaf,
        List<DivergenciaTotalesR1> divergencias)
    {
        var aseId = leaf.Ase.Id;
        var nombre = Normalizar(leaf.Ase.NombreCorto);

        var nameRow = filas.FirstOrDefault(r => string.Equals(Normalizar(Texto(celdas, "B" + r)), nombre, StringComparison.Ordinal));
        if (nameRow == 0)
        {
            throw Fail($"el gate no encontró el bloque del ASE {aseId} (B='{leaf.Ase.NombreCorto}') en la hoja '{HojaR1}'.");
        }

        var totOptRow = filas
            .Where(r => r > nameRow)
            .FirstOrDefault(r => Normalizar(Texto(celdas, "C" + r)) == "TOTAL" && Normalizar(Texto(celdas, "D" + r)) == "OPORTUNO");
        if (totOptRow == 0)
        {
            throw Fail($"el gate no encontró el visible TOT_OPT (C='TOTAL', D='OPORTUNO') del ASE {aseId} en '{HojaR1}'.");
        }

        var extempRow = filas
            .Where(r => r > totOptRow)
            .FirstOrDefault(r => Normalizar(Texto(celdas, "D" + r)) == "EXTEMPORANEO");
        if (extempRow == 0)
        {
            throw Fail($"el gate no encontró el visible EXTEMP (D='EXTEMPORANEO') del ASE {aseId} en '{HojaR1}'.");
        }

        // Filas Mes/Total del bloque (firma B='Mes' ∧ C='Total'): la geometría real del período.
        var mesRows = filas
            .Where(r => r > nameRow && r < totOptRow
                && Normalizar(Texto(celdas, "B" + r)) == "MES"
                && Normalizar(Texto(celdas, "C" + r)) == "TOTAL")
            .OrderBy(r => r)
            .ToList();
        if (mesRows.Count == 0)
        {
            throw Fail($"el gate no encontró filas 'Mes/Total' en el bloque del ASE {aseId} en '{HojaR1}' (se necesita el dato por firma).");
        }

        var totOptRef = "F" + totOptRow.ToString(CultureInfo.InvariantCulture);
        Comparar(divergencias, aseId, totOptRef, "TOT_OPT", leaf.R1.TotalOportunoEsperadoPorAse, EvaluarVisible(celdas, totOptRef));

        var extempRef = "F" + extempRow.ToString(CultureInfo.InvariantCulture);
        Comparar(divergencias, aseId, extempRef, "EXTEMP", leaf.R1.ExtemporaneoEsperadoPorAse, EvaluarVisible(celdas, extempRef));

        ValidarTdfG(celdas, aseId, "G" + totOptRow.ToString(CultureInfo.InvariantCulture), mesRows, divergencias);
    }

    /// <summary>
    /// Evalúa un visible R1: fórmula re-derivable contra literales del mismo workbook, o literal
    /// propio (0-Aplic explícito del contrato T2).
    /// </summary>
    private static decimal EvaluarVisible(IReadOnlyDictionary<string, CeldaR1Xlsx> celdas, string celdaRef)
    {
        var celda = celdas.GetValueOrDefault(celdaRef)
            ?? throw Fail($"{HojaR1}!{celdaRef}: no existe la celda del visible (se necesita el texto de fórmula o su literal).");

        if (celda.Formula is not null)
        {
            return EvaluarFormula(celdas, celdaRef, celda.Formula);
        }

        if (celda.Numero is not null)
        {
            return celda.Numero.Value;
        }

        throw Fail($"{HojaR1}!{celdaRef}: la celda no es fórmula ni literal numérico (se necesita un valor re-derivable).");
    }

    /// <summary>
    /// Evalúa una fórmula lineal resolviendo cada referencia contra su literal en el mismo workbook.
    /// Fail-fast si la fórmula no trae términos, o si una referencia no existe, es fórmula (caché
    /// stale) o no tiene literal numérico.
    /// </summary>
    private static decimal EvaluarFormula(IReadOnlyDictionary<string, CeldaR1Xlsx> celdas, string celdaVisible, string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            throw Fail($"{HojaR1}!{celdaVisible}: la celda es fórmula sin texto propio (se necesita el texto de fórmula).");
        }

        var terminos = TerminoFormula.Matches(formula);
        if (terminos.Count == 0)
        {
            throw Fail($"{HojaR1}!{celdaVisible}: la fórmula '{formula}' no tiene términos de celda reconocibles.");
        }

        var total = 0m;
        foreach (Match termino in terminos)
        {
            var referencia = NormalizarRef(termino.Groups["ref"].Value);
            var signo = termino.Groups["signo"].Value == "-" ? -1m : 1m;

            var celda = celdas.GetValueOrDefault(referencia)
                ?? throw Fail($"{HojaR1}!{celdaVisible}: la fórmula referencia '{referencia}', que no existe en el workbook (se necesita el literal de '{referencia}').");
            if (celda.Formula is not null)
            {
                throw Fail($"{HojaR1}!{celdaVisible}: la fórmula referencia '{referencia}', que es fórmula en el workbook (se necesita su literal; prohibido usar el caché <v>).");
            }

            if (celda.Numero is null)
            {
                throw Fail($"{HojaR1}!{celdaVisible}: la fórmula referencia '{referencia}', que no tiene literal numérico (se necesita el dato).");
            }

            total += signo * celda.Numero.Value;
        }

        return total;
    }

    /// <summary>
    /// Valida el total TDF (columna G): su conjunto de referencias debe ser exactamente el de las
    /// celdas G de las filas <c>Mes/Total</c> del bloque, excluida la última por orden físico
    /// (forma del manual, E2 congelado). Sin agregado de dominio (no existe para TDF).
    /// </summary>
    private static void ValidarTdfG(
        IReadOnlyDictionary<string, CeldaR1Xlsx> celdas,
        int aseId,
        string celdaG,
        IReadOnlyList<int> mesRows,
        List<DivergenciaTotalesR1> divergencias)
    {
        var celda = celdas.GetValueOrDefault(celdaG);
        if (celda?.Formula is null)
        {
            throw Fail($"{HojaR1}!{celdaG}: el visible del total TDF (ASE {aseId}) no es fórmula (se necesita el texto de fórmula).");
        }

        var reales = TerminoFormula.Matches(celda.Formula)
            .Select(m => NormalizarRef(m.Groups["ref"].Value))
            .ToList();

        var esperadas = (mesRows.Count > 1 ? mesRows.Take(mesRows.Count - 1) : mesRows)
            .Select(r => "G" + r.ToString(CultureInfo.InvariantCulture))
            .OrderBy(r => r, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var conjuntoEsperado = new HashSet<string>(esperadas, StringComparer.OrdinalIgnoreCase);
        var conjuntoReal = new HashSet<string>(reales, StringComparer.OrdinalIgnoreCase);
        if (!conjuntoReal.SetEquals(conjuntoEsperado))
        {
            divergencias.Add(new DivergenciaTotalesR1(
                aseId,
                celdaG,
                "TDF",
                string.Join("+", esperadas),
                string.Join("+", reales),
                "las referencias de la fórmula del total TDF no son las filas Mes/Total del bloque (forma E2: ΣG(Mes) − G(última Mes))"));
        }
    }

    private static void Comparar(
        List<DivergenciaTotalesR1> divergencias,
        int aseId,
        string celda,
        string etiqueta,
        decimal esperado,
        decimal evaluado)
    {
        if (Math.Abs(esperado - evaluado) <= Tolerancia)
        {
            return;
        }

        divergencias.Add(new DivergenciaTotalesR1(
            aseId,
            celda,
            etiqueta,
            Formatear(esperado),
            Formatear(evaluado),
            "el total visible del workbook no coincide con el dominio por firma (±0.5)"));
    }

    // ── Lectura ZIP+XML (BCL puro) ────────────────────────────────────────────────────────────

    /// <summary>
    /// Lee todas las celdas de una hoja del workbook como <c>ref → (fórmula, texto, número)</c>.
    /// Para una celda con fórmula NO se expone el caché <c>&lt;v&gt;</c> (prohibido por R-G-1).
    /// Público para que la evidencia E2 y la regresión puedan leer insumos reales con el mismo
    /// parser (sin Excel/COM).
    /// </summary>
    public static IReadOnlyDictionary<string, CeldaR1Xlsx> LeerCeldas(string rutaWorkbook, string hoja = HojaR1)
    {
        if (string.IsNullOrWhiteSpace(rutaWorkbook))
        {
            throw new ArgumentException("La ruta del workbook es requerida.", nameof(rutaWorkbook));
        }

        if (!File.Exists(rutaWorkbook))
        {
            throw new CalculoInvalidoException(
                CodigoError.FormatoFuente,
                $"No existe el workbook a validar: '{rutaWorkbook}'.");
        }

        using var zip = ZipFile.OpenRead(rutaWorkbook);
        var shared = LeerSharedStrings(zip);
        var rutaHoja = ResolverHoja(zip, hoja)
            ?? throw new CalculoInvalidoException(
                CodigoError.FormatoFuente,
                $"El workbook '{Path.GetFileName(rutaWorkbook)}' no trae la hoja '{hoja}'.");

        var doc = Cargar(zip, rutaHoja)
            ?? throw new CalculoInvalidoException(
                CodigoError.FormatoFuente,
                $"El workbook '{Path.GetFileName(rutaWorkbook)}' no tiene el XML de la hoja '{hoja}' ({rutaHoja}).");

        return LeerCeldasDeHoja(doc, shared);
    }

    private static Dictionary<string, CeldaR1Xlsx> LeerCeldasDeHoja(XmlDocument sheetXml, IReadOnlyList<string> shared)
    {
        var celdas = new Dictionary<string, CeldaR1Xlsx>(StringComparer.OrdinalIgnoreCase);
        var ns = new XmlNamespaceManager(sheetXml.NameTable);
        ns.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        foreach (XmlNode c in sheetXml.SelectNodes("//m:sheetData/m:row/m:c", ns)!)
        {
            var referencia = c.Attributes?["r"]?.Value;
            if (string.IsNullOrWhiteSpace(referencia))
            {
                continue;
            }

            var formulaNodo = c.SelectSingleNode("m:f", ns);
            var valorNodo = c.SelectSingleNode("m:v", ns);
            var tipo = c.Attributes?["t"]?.Value;

            string? formula = null;
            string? texto = null;
            decimal? numero = null;

            if (formulaNodo is not null)
            {
                formula = formulaNodo.InnerText;
            }
            else
            {
                // SOLO literales: nunca se lee el <v> de una celda con fórmula.
                if (tipo == "s")
                {
                    if (valorNodo is not null
                        && int.TryParse(valorNodo.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx)
                        && idx >= 0 && idx < shared.Count)
                    {
                        texto = shared[idx];
                    }
                }
                else if (tipo is "inlineStr")
                {
                    texto = c.SelectSingleNode("m:is", ns)?.InnerText ?? string.Empty;
                }
                else if (tipo is "str" or "b" or "e")
                {
                    texto = valorNodo?.InnerText;
                }
                else if (valorNodo is not null
                    && decimal.TryParse(valorNodo.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var n))
                {
                    numero = n;
                }
                else if (valorNodo is not null && !string.IsNullOrWhiteSpace(valorNodo.InnerText))
                {
                    texto = valorNodo.InnerText;
                }
            }

            if (formula is null && texto is null && numero is null)
            {
                continue;
            }

            celdas[referencia!] = new CeldaR1Xlsx(formula, texto, numero);
        }

        return celdas;
    }

    private static string? ResolverHoja(ZipArchive zip, string hoja)
    {
        var workbookXml = Cargar(zip, "xl/workbook.xml");
        if (workbookXml is null)
        {
            return null;
        }

        var relaciones = LeerRelaciones(zip);
        var ns = new XmlNamespaceManager(workbookXml.NameTable);
        ns.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        foreach (XmlNode sheet in workbookXml.SelectNodes("//m:sheets/m:sheet", ns)!)
        {
            var nombre = sheet.Attributes?["name"]?.Value;
            var rid = sheet.Attributes?["r:id"]?.Value;
            if (!string.Equals(nombre, hoja, StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrWhiteSpace(rid)
                || !relaciones.TryGetValue(rid!, out var target))
            {
                continue;
            }

            return target.StartsWith("/", StringComparison.Ordinal)
                ? target.TrimStart('/')
                : "xl/" + target;
        }

        return null;
    }

    private static List<string> LeerSharedStrings(ZipArchive zip)
    {
        var doc = Cargar(zip, "xl/sharedStrings.xml");
        var lista = new List<string>();
        if (doc is null)
        {
            return lista;
        }

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        foreach (XmlNode si in doc.SelectNodes("//m:si", ns)!)
        {
            lista.Add(si.InnerText);
        }

        return lista;
    }

    private static Dictionary<string, string> LeerRelaciones(ZipArchive zip)
    {
        var doc = Cargar(zip, "xl/_rels/workbook.xml.rels");
        var mapa = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (doc is null)
        {
            return mapa;
        }

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("p", "http://schemas.openxmlformats.org/package/2006/relationships");
        foreach (XmlNode rel in doc.SelectNodes("//p:Relationship", ns)!)
        {
            var id = rel.Attributes?["Id"]?.Value;
            var target = rel.Attributes?["Target"]?.Value;
            if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
            {
                mapa[id!] = target!;
            }
        }

        return mapa;
    }

    private static XmlDocument? Cargar(ZipArchive zip, string path)
    {
        var entry = zip.GetEntry(path);
        if (entry is null)
        {
            return null;
        }

        using var stream = entry.Open();
        var doc = new XmlDocument { XmlResolver = null };
        doc.Load(stream);
        return doc;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────────────────────

    private static IReadOnlyList<int> FilasOrdenadas(IReadOnlyDictionary<string, CeldaR1Xlsx> celdas) =>
        celdas.Keys
            .Select(r => ReferenciaCelda.Match(r))
            .Where(m => m.Success)
            .Select(m => int.Parse(m.Groups["fila"].Value, CultureInfo.InvariantCulture))
            .Distinct()
            .OrderBy(r => r)
            .ToList();

    private static string Texto(IReadOnlyDictionary<string, CeldaR1Xlsx> celdas, string celdaRef) =>
        celdas.GetValueOrDefault(celdaRef)?.Texto ?? string.Empty;

    private static string NormalizarRef(string token) =>
        token.Replace("$", string.Empty, StringComparison.Ordinal).ToUpperInvariant();

    private static string Formatear(decimal valor) => valor.ToString(CultureInfo.InvariantCulture);

    private static CalculoInvalidoException Fail(string mensaje) =>
        new(CodigoError.Validacion, $"Gate R1: {mensaje}");

    private static string Normalizar(string? texto)
    {
        if (string.IsNullOrWhiteSpace(texto))
        {
            return string.Empty;
        }

        var descompuesto = texto.Normalize(NormalizationForm.FormD);
        var constructor = new StringBuilder(descompuesto.Length);
        foreach (var caracter in descompuesto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(caracter) != UnicodeCategory.NonSpacingMark)
            {
                constructor.Append(char.ToUpperInvariant(caracter));
            }
        }

        return constructor.ToString().Trim();
    }
}

/// <summary>
/// Celda cruda de una hoja leída por <see cref="ValidadorTotalesR1Workbook.LeerCeldas"/>.
/// <see cref="Formula"/> no nulo ⇒ celda de fórmula (su <c>&lt;v&gt;</c> NO se expone).
/// </summary>
public sealed record CeldaR1Xlsx(string? Formula, string? Texto, decimal? Numero);

/// <summary>
/// Plan 32 (T1, pieza c): celda INTERIOR de <c>Reporte Componentes R1</c> congelada como objetivo del
/// gate-v2 <see cref="ValidadorTotalesR1Workbook.ExigirFormaRecompuestaInterior"/>. <see cref="Formula"/>
/// nulo = clase L-1 (el manual no trae <c>&lt;f&gt;</c>; cuando <see cref="ValorLiteral"/> no es nulo,
/// la salida debe traer ese literal). <paramref name="Etiqueta"/> es solo trazabilidad
/// (SUB_EMP/SUB_TDF/SUB_L/EXT_INT/SUBS/AFASEO).
/// </summary>
public sealed record InteriorR1Esperado(int AseId, string Celda, string Etiqueta, string? Formula, decimal? ValorLiteral = null);

/// <summary>Divergencia del gate R1 (celda visible vs dominio por firma).</summary>
public sealed record DivergenciaTotalesR1(int AseId, string Celda, string Etiqueta, string Esperado, string Evaluado, string Motivo);

/// <summary>
/// Resultado del gate R1: lista de divergencias (vacía = PASS) y mensaje legible que nombra
/// hoja + celda + esperado + evaluado.
/// </summary>
public sealed class ResultadoValidacionTotalesR1
{
    /// <summary>Crea el resultado con las divergencias detectadas.</summary>
    public ResultadoValidacionTotalesR1(IReadOnlyList<DivergenciaTotalesR1> divergencias)
    {
        Divergencias = divergencias;
    }

    /// <summary>Divergencias detectadas (vacía = el gate cierra).</summary>
    public IReadOnlyList<DivergenciaTotalesR1> Divergencias { get; }

    /// <summary>Verdad si el gate cierra (sin divergencias).</summary>
    public bool EsValido => Divergencias.Count == 0;

    /// <summary>Mensaje compacto del gate (nombra hoja + celda + esperado + evaluado por divergencia).</summary>
    public string Mensaje
    {
        get
        {
            if (EsValido)
            {
                return $"Gate R1 ({ValidadorTotalesR1Workbook.HojaR1}): sin divergencias.";
            }

            var lineas = Divergencias.Select(d =>
                $"ASE {d.AseId} {ValidadorTotalesR1Workbook.HojaR1}!{d.Celda} [{d.Etiqueta}]: esperado {d.Esperado}, evaluado {d.Evaluado} ({d.Motivo}).");
            return $"Gate R1 ({ValidadorTotalesR1Workbook.HojaR1}): {Divergencias.Count} divergencia(s):\n" + string.Join("\n", lineas);
        }
    }
}
