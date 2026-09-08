namespace Remuneracion.Core.Models;

/// <summary>
/// Inputs 2.2 (HU-08) de una hoja <c>Recaudo *</c>: valores por concepto/ASE (zona de datos
/// filas ~3–28, docx V9). Fuente: <c>Consolidado/Conciliaciones/Conjunta {prefijo}*.xlsx</c>
/// hoja <c>RESUMEN MES</c> (T0-0.6 cerrado con evidencia: alineación 1:1 fila por fila).
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
    /// Celdas de la hoja <c>Recaudo *</c> (ref → valor). Incluye D3:D9 (OPORTUNO),
    /// D12:D18 (EXTEMP), D21:D27 (TOTAL) y sus E (número de registros).
    /// </summary>
    public IReadOnlyDictionary<string, decimal> Celdas { get; set; } = new Dictionary<string, decimal>();

    /// <summary>
    /// Total OPORTUNO de la hoja (D9) para el resumen por empresa del log/Serilog.
    /// </summary>
    public decimal TotalOportuno { get; set; }

    /// <summary>
    /// Total EXTEMPORÁNEO de la hoja (D18).
    /// </summary>
    public decimal TotalExtemporaneo { get; set; }

    /// <summary>
    /// Gran total de la hoja (D27).
    /// </summary>
    public decimal Total { get; set; }
}