namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 61 (H3.1 / S3.1, §2-D-H3): UNA novedad de datos de la corrida, en lenguaje de NEGOCIO.
/// Es lo que alimenta la ficha del reporte visible para el administrativo, así que sus textos
/// (<see cref="Reporte"/>, <see cref="ConceptoOCelda"/>, <see cref="Dueno"/>, <see cref="QueRevisar"/>,
/// <see cref="Explicacion"/>) están PROHIBIDOS de jerga: sin <c>ERR-*</c>, sin <c>&lt;f&gt;</c> y sin
/// nombres de clase (G-TRAZA / D-H4).
///
/// <para><b>Criterio de clase (no de caso).</b> Toda divergencia de DATOS con workbook escribible es
/// novedad: <c>Σ</c> visibles-empresa vs bloque, DetRetri calculado vs R10, validación cruzada 2.7
/// no declarada. Toda imposibilidad de construir el artefacto sigue siendo fallo duro y NO produce
/// novedad.</para>
///
/// <para><b>Montos.</b> Son <c>decimal?</c> a propósito: cuando un gate no expone un número
/// comparable (p. ej. una etiqueta de fuente no reconocida) el modelo NO inventa un cero
/// (<c>0</c> sería un dato falso). El reporte omite el renglón que no aplica.</para>
/// </summary>
/// <param name="Reporte">Nombre de negocio del reporte/fuente que se revisó (p. ej. "R10 del período").</param>
/// <param name="Hoja">Hoja del Excel implicada, con su nombre real (p. ej. <c>DetRetri2026091</c>).</param>
/// <param name="AseId">ASE afectado (1..5); <c>0</c> = afecta al período completo (Σ de los cinco).</param>
/// <param name="ConceptoOCelda">Concepto o celda que no cuadra, nombrado por su ETIQUETA.</param>
/// <param name="ValorObservado">Lo que muestra el archivo/fuente; <c>null</c> si el gate no lo expone.</param>
/// <param name="ValorEsperado">Lo que esperaba la validación; <c>null</c> si el gate no lo expone.</param>
/// <param name="Delta">Diferencia observada (observado − esperado); <c>null</c> si no aplica.</param>
/// <param name="Dueno">De quién es la decisión (negocio), nunca un nombre de clase.</param>
/// <param name="QueRevisar">Instrucción concreta de revisión, por ETIQUETA contra el administrativo y el R10.</param>
/// <param name="Explicacion">Lectura de negocio de qué pasó (contexto, no jerga).</param>
public sealed record NovedadEntrega(
    string Reporte,
    string Hoja,
    int AseId,
    string ConceptoOCelda,
    decimal? ValorObservado,
    decimal? ValorEsperado,
    decimal? Delta,
    string Dueno,
    string QueRevisar,
    string? Explicacion = null);
