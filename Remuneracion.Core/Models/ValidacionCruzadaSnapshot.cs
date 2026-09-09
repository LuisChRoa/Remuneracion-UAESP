namespace Remuneracion.Core.Models;

/// <summary>
/// HU-13 (2.7): snapshot-oráculo de la validación por empresa
/// (hojas <c>VALIDACION_*</c>). Congelado por T0 (Plan 13 §4 Fase 0) contra ambos canónicos:
/// las filas 3..7 de cada hoja <c>VALIDACION_*</c> son los ASE 1..5 (columna B numérica);
/// la columna O = H−N (Recaudo * vs REMUNERACION_*) y la columna P = INT(O)=0 (booleano,
/// caché golden TRUE cuando la diferencia cierra).
/// </summary>
public sealed class ValidacionEmpresaSnapshot
{
    /// <summary>
    /// Nombre de la empresa del catálogo <see cref="EmpresaFacturacion.Nombre"/>
    /// (p. ej. "Enel").
    /// </summary>
    public string Empresa { get; set; } = string.Empty;

    /// <summary>
    /// Id del ASE (1..5) al que corresponde la fila leída de la hoja VALIDACION_*.
    /// </summary>
    public int AseId { get; set; }

    /// <summary>
    /// Valor cacheado de la columna O (H−N). Golden Q1/Q2 = 0. Gate: |O| ≤ ±0.5.
    /// </summary>
    public decimal DiferenciaO { get; set; }

    /// <summary>
    /// Valor cacheado de la columna P (INT(O)=0). Golden Q1/Q2 = TRUE. Gate: P == true exacto.
    /// </summary>
    public bool VerificacionP { get; set; }
}

/// <summary>
/// HU-13 (2.7): snapshot-oráculo <c>DetValiRetri</c> de UN ASE (columnas D).
/// Congelado por T0: D16..D20 = D9..D13 − CONSOLIDADO U104..U108 (diferencias float ≤ ±0.5 en
/// los goldens) y D24..D28 = booleanos de composición TRUE. La fila 21 (Total) DIVERGE en ambos
/// canónicos (Q1 D21 = −0.62, Q2 D21 = −1.71) y queda EXCLUIDA del gate como divergencia
/// documentada (D6); lo mismo para D9:D14 (valores fuente ≠ ROUND, hallazgo HU-12).
/// </summary>
public sealed class DetValiRetriSnapshot
{
    /// <summary>
    /// ASE al que corresponde el snapshot (matcheo estricto por <see cref="Ase.Id"/>).
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Diferencias por ASE leídas del caché (D16..D20 = filas 15+aseId). Celda → valor.
    /// Gate: cada |Diferencia| ≤ ±0.5 (error que nombra ASE + celda).
    /// </summary>
    public IReadOnlyList<DetValiRetriCeldaValor> DiferenciasAse { get; set; } = [];

    /// <summary>
    /// Verificaciones booleanas de composición por ASE (D24..D28 = filas 23+aseId).
    /// Celda → resultado booleano cacheado. Gate: == true exacto (error que nombra ASE + celda).
    /// </summary>
    public IReadOnlyList<DetValiRetriCeldaVerificacion> VerificacionesAse { get; set; } = [];

    /// <summary>
    /// Diferencia de la fila Total (D21 = D14 − CONSOLIDADO U109). Documentada: en los goldens
    /// Q1/Q2 NO cierra ±0.5 (acumulado de ruido float de las 5 filas). NUNCA gate de error.
    /// </summary>
    public decimal DiferenciaTotalD21 { get; set; }

    /// <summary>
    /// Verificación booleana de la fila TOTAL (D29). Golden Q1/Q2 = TRUE. Gate: == true exacto.
    /// </summary>
    public bool VerificacionTotalD29 { get; set; }
}

/// <summary>
/// Valor leído de una celda de diferencia DetValiRetri (caché del workbook).
/// </summary>
public sealed class DetValiRetriCeldaValor
{
    /// <summary>Referencia de la celda (p. ej. "D16").</summary>
    public string Celda { get; set; } = string.Empty;

    /// <summary>Valor numérico cacheado.</summary>
    public decimal Valor { get; set; }
}

/// <summary>
/// Valor booleano leído de una celda de verificación DetValiRetri (caché del workbook).
/// </summary>
public sealed class DetValiRetriCeldaVerificacion
{
    /// <summary>Referencia de la celda (p. ej. "D24").</summary>
    public string Celda { get; set; } = string.Empty;

    /// <summary>Resultado booleano cacheado (Excel booleano t="b": 1 = TRUE).</summary>
    public bool Verificacion { get; set; }
}

/// <summary>
/// HU-13 (2.7): snapshot-oráculo de validaciones cruzadas de UN ASE en UN período.
/// Oráculo de LECTURA (D1): lo puebla <c>IValidacionOracleReader</c> (Infrastructure, OpenXML
/// read-only, asserts W2) y lo consume el validador de dominio — que NUNCA abre .xlsx.
/// Snapshot ausente/vacío para un ASE = comportamiento HU-12 puro (gates 2.7 no corren).
/// </summary>
public sealed class ValidacionCruzadaSnapshot
{
    /// <summary>
    /// ASE del snapshot (matcheo estricto por <see cref="Ase.Id"/>).
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Validación por empresa (Recaudo * vs REMUNERACION_*) de la fila de ESTE ASE
    /// (5 empresas; O/P en la fila 2+aseId de cada hoja VALIDACION_*).
    /// </summary>
    public IReadOnlyList<ValidacionEmpresaSnapshot> PorEmpresa { get; set; } = [];

    /// <summary>
    /// Snapshot DetValiRetri de este ASE. <c>null</c> = sin gate DetValiRetri para el ASE.
    /// </summary>
    public DetValiRetriSnapshot? DetValiRetri { get; set; }

    /// <summary>
    /// <c>VALIDACION_TOTAL</c>: valor cacheado de la fila TOTAL O9 (= H9−N9 = Σ diferencias).
    /// Golden Q1/Q2 = 0. Gate: |ValidacionTotal| ≤ ±0.5.
    /// </summary>
    public decimal ValidacionTotal { get; set; }

    /// <summary>
    /// <c>VALIDACION_TOTAL</c>: valor cacheado de la fila TOTAL P9 (INT(O9)=0).
    /// Golden Q1/Q2 = TRUE. Gate: == true exacto.
    /// </summary>
    public bool ValidacionTotalOkP { get; set; }

    /// <summary>
    /// Controles <c>Valida -*</c> leídos (celda → valor caché). T0 (0.5): las hojas
    /// Valida -Remunera/-Anticipos/-Control Recaudo son vistas lado-a-lado SIN celda de
    /// diferencia/booleana estable → sin gate numérico; el diccionario es informativo/log.
    /// </summary>
    public IReadOnlyDictionary<string, decimal> Controles { get; set; } = new Dictionary<string, decimal>();

    /// <summary>
    /// HU-14 (W-1): sub-bloques booleanos de <c>VALIDACION_TOTAL</c> (C15/D25/O25/D34/F34)
    /// leídos del caché (celda → booleano). Golden Q1/Q2 = TRUE (amparo T0-0.4 HU-13).
    /// Gate: == true exacto por ASE con error que nombra ASE · VALIDACION_TOTAL · celda.
    /// Vacío = HU-13 puro (campo aditivo, compatibilidad).
    /// </summary>
    public IReadOnlyDictionary<string, bool> SubBloquesValidacionTotal { get; set; } = new Dictionary<string, bool>();
}
