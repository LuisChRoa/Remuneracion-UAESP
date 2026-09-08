namespace Remuneracion.Core.Models;

/// <summary>
/// Catálogo 2.2 de las empresas de facturación (HU-08). Nombres EXACTOS del template:
/// hojas <c>REMUNERACION_*</c>, hojas <c>Recaudo *</c> y labels de las filas detalle
/// R1/R2/R4 (<see cref="LabelTemplate"/>). El prefijo de conciliación localiza el archivo
/// <c>Consolidado/Conciliaciones/Conjunta {prefijo}*.xlsx</c> (T0-0.6: fuente de las hojas
/// <c>Recaudo *</c>).
/// </summary>
public sealed class EmpresaFacturacion
{
    public int Id { get; set; }

    /// <summary>
    /// Nombre corto de la empresa (clave estable del catálogo).
    /// </summary>
    public string Nombre { get; set; } = string.Empty;

    /// <summary>
    /// Label exacto de las filas detalle por empresa en R1/R2/R4 del template
    /// (p. ej. "ENEL", "OCCIDENTE", "RECIPROCIDAD", "ENERBIT", "CIUDAD LIMPIA-ACUEDUCTO").
    /// </summary>
    public string LabelTemplate { get; set; } = string.Empty;

    /// <summary>
    /// Nombre exacto de la hoja <c>REMUNERACION_*</c> (100 % fórmulas; se protege, no se escribe).
    /// </summary>
    public string HojaRemuneracion { get; set; } = string.Empty;

    /// <summary>
    /// Nombre exacto de la hoja <c>Recaudo *</c> (se escribe en valores, filas ~3–28).
    /// </summary>
    public string HojaRecaudo { get; set; } = string.Empty;

    /// <summary>
    /// Prefijo del archivo de conciliación en <c>Consolidado/Conciliaciones/</c>
    /// (p. ej. "Conjunta ENEL", "Directa", "Conjunta Otros").
    /// </summary>
    public string PrefijoConciliacion { get; set; } = string.Empty;

    /// <summary>
    /// Nombre exacto de la hoja <c>VALIDACION_*</c> (fórmulas; se protege, no se implementa — 2.7).
    /// </summary>
    public string HojaValidacion { get; set; } = string.Empty;

    /// <summary>
    /// Nombre exacto de la hoja <c>GERENTES_*</c> (fórmulas puras; se protege).
    /// </summary>
    public string HojaGerentes { get; set; } = string.Empty;

    /// <summary>
    /// Catálogo 2.2 fijo: las 5 empresas de facturación del template (V1 del plan).
    /// </summary>
    public static IReadOnlyList<EmpresaFacturacion> Catalogo { get; } =
    [
        new()
        {
            Id = 1,
            Nombre = "ReciprocidadEaab",
            LabelTemplate = "RECIPROCIDAD",
            HojaRemuneracion = "REMUNERACION Reciprocidad EAAB",
            HojaRecaudo = "Recaudo EAAB Reciprocidad",
            PrefijoConciliacion = "Conjunta Recip",
            HojaValidacion = "VALIDACION_RECIP",
            HojaGerentes = "GERENTES_Reciprocidad_ EAAB"
        },
        new()
        {
            Id = 2,
            Nombre = "Enel",
            LabelTemplate = "ENEL",
            HojaRemuneracion = "REMUNERACION_ENEL",
            HojaRecaudo = "Recaudo ENEL",
            PrefijoConciliacion = "Conjunta ENEL",
            HojaValidacion = "VALIDACION_ENEL",
            HojaGerentes = "GERENTES_ENEL"
        },
        new()
        {
            Id = 3,
            Nombre = "Enerbit",
            LabelTemplate = "ENERBIT",
            HojaRemuneracion = "REMUNERACION_ENERBIT",
            HojaRecaudo = "Recaudo ENERBIT",
            PrefijoConciliacion = "Conjunta ENERBIT",
            HojaValidacion = "VALIDACION_ENERBIT",
            HojaGerentes = "GERENTES_ENERBIT"
        },
        new()
        {
            Id = 4,
            Nombre = "OccidenteDirecta",
            LabelTemplate = "OCCIDENTE",
            HojaRemuneracion = "REMUNERACION_OCCIDENTE_ Directa",
            HojaRecaudo = "Recaudo Directa Occidente",
            PrefijoConciliacion = "Directa",
            HojaValidacion = "VALIDACION_OCCIDENTE",
            HojaGerentes = "GERENTES_OCCIDENTE _ Directa"
        },
        new()
        {
            Id = 5,
            Nombre = "EaabCiudadLimpia",
            LabelTemplate = "CIUDAD LIMPIA-ACUEDUCTO",
            HojaRemuneracion = "REMUNERACION_EAAB-CL",
            HojaRecaudo = "Recaudo EAAB + Ciud Limp",
            PrefijoConciliacion = "Conjunta Otros",
            HojaValidacion = "VALIDACION_EAAB-CL",
            HojaGerentes = "GERENTES_EAAB-CL"
        }
    ];

    /// <summary>
    /// Obtiene la empresa del catálogo por Id (1..5).
    /// </summary>
    public static EmpresaFacturacion Obtener(int id) =>
        Catalogo.FirstOrDefault(e => e.Id == id)
        ?? throw new ArgumentOutOfRangeException(nameof(id), $"No existe la empresa de facturación {id}.");
}