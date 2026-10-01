namespace Remuneracion.Core.Models;

/// <summary>
/// Inputs 2.2 (HU-08) de una hoja <c>Recaudo *</c>: valores por concepto/ASE (zona de datos
/// filas ~3–28, docx V9). Fuente: <c>{periodo}/Conciliaciones/Conjunta {prefijo}*.xlsx</c>
/// hoja <c>RESUMEN MES</c> (HU-20/G1-D1). Hu-20/G2-D2: Q1 lee D/E (VALOR 1°Q) y Q2 lee F/G
/// (VALOR 2°Q) según <see cref="Periodo.NumeroQuincena"/>.
/// La fila 29+ (validaciones =SUM) va al mapa de fórmulas protegidas.
/// </summary>
public sealed class RecaudoEmpresaInputs
{
    /// <summary>
    /// Empresa de facturación (catálogo 2.2).
    /// </summary>
    public EmpresaFacturacion Empresa { get; set; } = new();

    /// <summary>
    /// Nombre exacto de la hoja <c>Recaudo *</c> destino.
    /// </summary>
    public string HojaRecaudo { get; set; } = string.Empty;

    /// <summary>
    /// Celdas de la hoja <c>Recaudo *</c> (ref → valor). Incluye el bloque OPORTUNO (filas 3–9),
    /// EXTEMP (12–18) y TOTAL (21–27) con sus pares valor/n° registros. Las letras de columna
    /// reflejan el destino REAL de la quincena: D/E en Q1, F/G en Q2 (HU-20/G2-D2).
    /// </summary>
    public IReadOnlyDictionary<string, decimal> Celdas { get; set; } = new Dictionary<string, decimal>();

    /// <summary>
    /// Total OPORTUNO de la hoja (fila 9 de la columna de valor: D9 en Q1, F9 en Q2).
    /// </summary>
    public decimal TotalOportuno { get; set; }

    /// <summary>
    /// Total EXTEMPORÁNEO de la hoja (fila 18 de la columna de valor: D18 en Q1, F18 en Q2).
    /// </summary>
    public decimal TotalExtemporaneo { get; set; }

    /// <summary>
    /// Gran total de la hoja (fila 27 de la columna de valor: D27 en Q1, F27 en Q2).
    /// </summary>
    public decimal Total { get; set; }
}
