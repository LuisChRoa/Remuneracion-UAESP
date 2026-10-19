using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 40 (T1, R-E-1/R-E-2 — D-A/D-D): resolutor ACOTADO del valor de una celda del workbook de
/// salida, usado por el gate W-1 para EVALUAR los 5 sub-bloques de <c>VALIDACION_TOTAL</c> desde
/// las celdas efectivamente escritas.
///
/// NO es un parser de Excel general. Solo reconoce las formas tabuladas por el T0/§2.2:
///   - una REFERENCIA de celda (<c>A1</c>, <c>'Hoja'!A1</c>);
///   - <c>SUM(rango)</c>;
///   - una suma/resta de referencias o de SUM (signos +/-);
///   - un literal numérico.
/// Cualquier otra forma → NO computable (el gate no evalúa en silencio). Los sub-bloques son
/// igualdades (<c>LHS=RHS</c>) y el reader parte el texto antes de llamar aquí.
///
/// Bottom-out: una celda SIN fórmula se lee de su <c>&lt;v&gt;</c> (literal, la app la escribe);
/// una celda CON fórmula se RESUELVE recursivamente por su texto — JAMÁS se usa el caché
/// <c>&lt;v&gt;</c> de una celda de fórmula como oráculo (D-A). Las fórmulas compartidas
/// (follower → maestro) se resuelven desplazando las referencias relativas del maestro.
/// </summary>
internal sealed class ResolutorValorValidacion
{
    private static readonly Regex CeldaRegex = new(
        @"(?<hoja>'(?<h1>[^']+)'|(?<h2>[A-Za-z0-9_. áéíóúÁÉÍÓÚñÑ()\-]+))!(?<ref>\$?[A-Z]{1,3}\$?\d+)|(?<ref2>\$?[A-Z]{1,3}\$?\d+)",
        RegexOptions.Compiled);

    private static readonly Regex RangoRegex = new(
        @"^\$?(?<c1>[A-Z]{1,3})\$?(?<r1>\d+):\$?(?<c2>[A-Z]{1,3})\$?(?<r2>\d+)$",
        RegexOptions.Compiled);

    private static readonly Regex SumRegex = new(
        @"^SUM\((?<rango>[^()]+)\)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex NumeroRegex = new(
        @"^[-+]?\d+(\.\d+)?([eE][-+]?\d+)?$",
        RegexOptions.Compiled);

    private readonly WorkbookPart _workbookPart;
    private readonly Dictionary<string, Worksheet?> _hojas = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Dictionary<string, Cell>> _celdas = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Dictionary<string, (string Texto, string Ancla)>> _maestros = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<(string Hoja, string Celda), decimal> _memo = [];
    private readonly HashSet<(string Hoja, string Celda)> _enCurso = [];

    public ResolutorValorValidacion(WorkbookPart workbookPart)
    {
        _workbookPart = workbookPart ?? throw new ArgumentNullException(nameof(workbookPart));
    }

    /// <summary>
    /// Evalúa una EXPRESIÓN de valor (ref / SUM(rango) / suma-resta / literal) en el contexto de
    /// la hoja indicada. <c>false</c> = forma no soportada o referencia irresoluble (no computable).
    /// </summary>
    public bool TryEvaluar(string hojaActual, string expresion, out decimal valor)
    {
        valor = 0m;
        var terminos = SepararTerminos(expresion);
        if (terminos.Count == 0)
        {
            return false;
        }

        var total = 0m;
        foreach (var (signo, termino) in terminos)
        {
            if (!TryEvaluarTermino(hojaActual, termino, out var parcial))
            {
                return false;
            }

            total += signo * parcial;
        }

        valor = total;
        return true;
    }

    private bool TryEvaluarTermino(string hojaActual, string termino, out decimal valor)
    {
        valor = 0m;
        termino = termino.Trim();
        if (termino.Length == 0)
        {
            return false;
        }

        if (NumeroRegex.IsMatch(termino))
        {
            return decimal.TryParse(termino, NumberStyles.Any, CultureInfo.InvariantCulture, out valor);
        }

        var sum = SumRegex.Match(termino);
        if (sum.Success)
        {
            return TrySumarRango(hojaActual, sum.Groups["rango"].Value.Trim(), out valor);
        }

        if (RangoRegex.IsMatch(termino))
        {
            return TrySumarRango(hojaActual, termino, out valor);
        }

        if (TryParsearReferencia(termino, hojaActual, out var hoja, out var celda))
        {
            return TryResolverCelda(hoja, celda, out valor);
        }

        return false;
    }

    private bool TrySumarRango(string hojaActual, string rango, out decimal valor)
    {
        valor = 0m;
        var match = RangoRegex.Match(rango);
        if (!match.Success)
        {
            // Plan 48 (H1.2): la v2 trae `E10 = SUM(D10)` — SUM de una CELDA ÚNICA (no rango).
            // Es la misma forma tabulada (`SUM(...)`) con un solo referente; sin esto la cadena
            // C14→F10→E10 no sería evaluable y el gate honesto fallaría (W-1/E1-1).
            return TryParsearReferencia(rango, hojaActual, out var hojaUnica, out var celdaUnica)
                && TryResolverCelda(hojaUnica, celdaUnica, out valor);
        }

        var col1 = ColumnaANumero(match.Groups["c1"].Value);
        var col2 = ColumnaANumero(match.Groups["c2"].Value);
        var fila1 = int.Parse(match.Groups["r1"].Value, CultureInfo.InvariantCulture);
        var fila2 = int.Parse(match.Groups["r2"].Value, CultureInfo.InvariantCulture);
        if (col1 > col2)
        {
            (col1, col2) = (col2, col1);
        }

        if (fila1 > fila2)
        {
            (fila1, fila2) = (fila2, fila1);
        }

        var total = 0m;
        for (var fila = fila1; fila <= fila2; fila++)
        {
            for (var col = col1; col <= col2; col++)
            {
                if (!TryResolverCelda(hojaActual, $"{NumeroAColumna(col)}{fila}", out var v))
                {
                    return false;
                }

                total += v;
            }
        }

        valor = total;
        return true;
    }

    private bool TryResolverCelda(string hoja, string celda, out decimal valor)
    {
        valor = 0m;
        var clave = (hoja, celda);
        if (_memo.TryGetValue(clave, out valor))
        {
            return true;
        }

        if (!_enCurso.Add(clave))
        {
            return false; // ciclo
        }

        try
        {
            var cell = ObtenerCelda(hoja, celda);
            if (cell is null)
            {
                valor = 0m; // celda ausente = 0 (semántica de Excel)
                _memo[clave] = valor;
                return true;
            }

            if (cell.CellFormula is null)
            {
                valor = LeerLiteral(cell);
                _memo[clave] = valor;
                return true;
            }

            var texto = TextoFormula(cell, hoja, celda);
            if (string.IsNullOrWhiteSpace(texto))
            {
                return false; // fórmula sin texto ni maestro compartido resoluble
            }

            if (!TryEvaluar(hoja, texto, out valor))
            {
                return false;
            }

            _memo[clave] = valor;
            return true;
        }
        finally
        {
            _enCurso.Remove(clave);
        }
    }

    /// <summary>
    /// Texto-<c>&lt;f&gt;</c> EFECTIVO de una celda: su propio texto, o el del maestro compartido
    /// desplazado a esta celda (follower). Vacío si no es fórmula o no hay maestro resoluble.
    /// </summary>
    public string TextoFormulaEfectiva(string hoja, string celda)
    {
        var cell = ObtenerCelda(hoja, celda);
        if (cell is null || cell.CellFormula is null)
        {
            return string.Empty;
        }

        return TextoFormula(cell, hoja, celda);
    }

    private string TextoFormula(Cell cell, string hoja, string celda)
    {
        var texto = cell.CellFormula?.Text ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(texto))
        {
            return texto;
        }

        // Follower de una fórmula compartida: maestro + desplazamiento de referencias relativas.
        var si = cell.CellFormula?.SharedIndex?.Value;
        if (si is null)
        {
            return string.Empty;
        }

        var maestro = ObtenerMaestroCompartido(hoja, si.Value.ToString(CultureInfo.InvariantCulture));
        if (maestro is null)
        {
            return string.Empty;
        }

        return DesplazarFormula(maestro.Value.Texto, maestro.Value.Ancla, celda);
    }

    private (string Texto, string Ancla)? ObtenerMaestroCompartido(string hoja, string si)
    {
        var indice = IndiceMaestros(hoja);
        return indice.TryGetValue(si, out var maestro) ? maestro : null;
    }

    private Dictionary<string, (string Texto, string Ancla)> IndiceMaestros(string hoja)
    {
        if (_maestros.TryGetValue(hoja, out var existente))
        {
            return existente;
        }

        var indice = new Dictionary<string, (string, string)>(StringComparer.Ordinal);
        var worksheet = ObtenerHoja(hoja);
        if (worksheet is not null)
        {
            foreach (var cell in worksheet.Descendants<Cell>())
            {
                var formula = cell.CellFormula;
                if (formula is null || formula.FormulaType?.Value != CellFormulaValues.Shared)
                {
                    continue;
                }

                var refMaestro = formula.Reference?.Value;
                var texto = formula.Text;
                if (string.IsNullOrEmpty(refMaestro) || string.IsNullOrWhiteSpace(texto))
                {
                    continue;
                }

                var ancla = refMaestro.Split(':')[0];
                var si = formula.SharedIndex?.Value.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
                indice[si] = (texto, ancla);
            }
        }

        _maestros[hoja] = indice;
        return indice;
    }

    private static string DesplazarFormula(string texto, string ancla, string destino)
    {
        if (!TryParsearRefSimple(ancla, out var anclaCol, out var anclaFila)
            || !TryParsearRefSimple(destino, out var destCol, out var destFila))
        {
            return texto;
        }

        var deltaCol = destCol - anclaCol;
        var deltaFila = destFila - anclaFila;

        return Regex.Replace(texto, @"(?<abs1>\$?)(?<col>[A-Z]{1,3})(?<abs2>\$?)(?<fila>\d+)", m =>
        {
            var col = m.Groups["col"].Value;
            var fila = int.Parse(m.Groups["fila"].Value, CultureInfo.InvariantCulture);
            if (m.Groups["abs1"].Value != "$")
            {
                col = NumeroAColumna(ColumnaANumero(col) + deltaCol);
            }

            if (m.Groups["abs2"].Value != "$")
            {
                fila += deltaFila;
            }

            return $"{m.Groups["abs1"].Value}{col}{m.Groups["abs2"].Value}{fila}";
        });
    }

    private bool TryParsearReferencia(string texto, string hojaActual, out string hoja, out string celda)
    {
        hoja = hojaActual;
        celda = string.Empty;
        var match = CeldaRegex.Match(texto);
        if (!match.Success)
        {
            return false;
        }

        if (match.Groups["ref"].Success)
        {
            hoja = match.Groups["h1"].Success ? match.Groups["h1"].Value : match.Groups["h2"].Value;
            celda = match.Groups["ref"].Value;
        }
        else
        {
            celda = match.Groups["ref2"].Value;
        }

        return true;
    }

    private static bool TryParsearRefSimple(string referencia, out int columna, out int fila)
    {
        columna = 0;
        fila = 0;
        var match = Regex.Match(referencia, @"^\$?(?<col>[A-Z]{1,3})\$?(?<fila>\d+)$");
        if (!match.Success)
        {
            return false;
        }

        columna = ColumnaANumero(match.Groups["col"].Value);
        fila = int.Parse(match.Groups["fila"].Value, CultureInfo.InvariantCulture);
        return true;
    }

    private Cell? ObtenerCelda(string hoja, string celda)
    {
        var mapa = IndiceCeldas(hoja);
        return mapa.TryGetValue(celda, out var cell) ? cell : null;
    }

    private Dictionary<string, Cell> IndiceCeldas(string hoja)
    {
        if (_celdas.TryGetValue(hoja, out var existente))
        {
            return existente;
        }

        var mapa = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        var worksheet = ObtenerHoja(hoja);
        if (worksheet is not null)
        {
            foreach (var cell in worksheet.Descendants<Cell>())
            {
                var referencia = cell.CellReference?.Value;
                if (!string.IsNullOrEmpty(referencia))
                {
                    mapa[referencia] = cell;
                }
            }
        }

        _celdas[hoja] = mapa;
        return mapa;
    }

    private Worksheet? ObtenerHoja(string nombre)
    {
        if (_hojas.TryGetValue(nombre, out var existente))
        {
            return existente;
        }

        Worksheet? worksheet = null;
        var sheet = _workbookPart.Workbook?.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is { } id && _workbookPart.GetPartById(id) is WorksheetPart parte)
        {
            worksheet = parte.Worksheet;
        }

        _hojas[nombre] = worksheet;
        return worksheet;
    }

    private static decimal LeerLiteral(Cell cell)
    {
        var texto = cell.CellValue?.InnerText;
        if (string.IsNullOrWhiteSpace(texto))
        {
            return 0m;
        }

        if (decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
        {
            return valor;
        }

        // Un literal booleano/string no aporta aritmética: 0 (no es una celda de valor de la cadena).
        return 0m;
    }

    /// <summary>
    /// Separa una expresión en términos con su signo. Prefija '+' si no empieza con signo; corta
    /// en '+'/'-' a profundidad 0 de paréntesis.
    /// </summary>
    private static List<(int Signo, string Termino)> SepararTerminos(string expresion)
    {
        var terminos = new List<(int, string)>();
        var s = (expresion ?? string.Empty).Trim();
        if (s.Length == 0)
        {
            return terminos;
        }

        if (s[0] != '+' && s[0] != '-')
        {
            s = "+" + s;
        }

        var profundidad = 0;
        var enComillas = false;
        var actual = new StringBuilder();
        foreach (var ch in s)
        {
            if (ch == '\'')
            {
                enComillas = !enComillas;
            }

            if (!enComillas)
            {
                if (ch == '(')
                {
                    profundidad++;
                }
                else if (ch == ')')
                {
                    profundidad--;
                }
            }

            if (!enComillas && (ch == '+' || ch == '-') && profundidad == 0 && actual.Length > 0)
            {
                AgregarTermino(terminos, actual.ToString());
                actual.Clear();
            }

            actual.Append(ch);
        }

        if (actual.Length > 0)
        {
            AgregarTermino(terminos, actual.ToString());
        }

        return terminos;
    }

    private static void AgregarTermino(List<(int, string)> terminos, string parte)
    {
        var signo = parte[0] == '-' ? -1 : 1;
        var termino = parte[1..].Trim();
        if (termino.Length > 0)
        {
            terminos.Add((signo, termino));
        }
    }

    private static int ColumnaANumero(string columna)
    {
        var numero = 0;
        foreach (var ch in columna.ToUpperInvariant())
        {
            numero = (numero * 26) + (ch - 'A' + 1);
        }

        return numero;
    }

    private static string NumeroAColumna(int numero)
    {
        var texto = string.Empty;
        while (numero > 0)
        {
            var resto = (numero - 1) % 26;
            texto = (char)('A' + resto) + texto;
            numero = (numero - 1) / 26;
        }

        return texto;
    }
}
