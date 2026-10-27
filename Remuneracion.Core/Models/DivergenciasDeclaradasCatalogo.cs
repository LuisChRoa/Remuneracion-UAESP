namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 42 (T1, D-C — patrón Plan 34 <see cref="SinonimosEmpresaR1"/>) + Plan 46 (H2, decisión del
/// Ingeniero 2026-10-17): catálogo de DIVERGENCIAS DECLARADAS como DATOS citados, más el matcher puro
/// <see cref="Buscar(IReadOnlyList{Entrada}, string, decimal, decimal)"/>.
///
/// Una divergencia declarada es una diferencia EXTERNA por construcción (discrepancia
/// oficial-vs-manual con dueño fuera de la aplicación). NO es un defecto interno: los defectos
/// de lectura, cálculo, colocación, sellado, naming, preflight o caché NUNCA son declarables
/// (taxonomía T0 del Plan 42). El catálogo es ABIERTO pero DISCIPLINADO: cada entrada exige sus
/// 6 campos citados (celda(s), origen del monto, dueño, cita de evidencia, fecha y condición de
/// retiro); sin ellos la entrada no aplica (jamás un verde silencioso). Añadir una divergencia
/// futura = añadir una FILA citada, nunca una rama de código.
///
/// <para>
/// <b>MECANISMO DE RETIRO (Plan 46 H2):</b> una entrada puede quedar <see cref="Entrada.Retirada"/>
/// (estado de retiro con <see cref="Entrada.FechaRetiro"/>, <see cref="Entrada.MotivoRetiro"/> y
/// <see cref="Entrada.EvidenciaRetiro"/> citados). Una entrada retirada YA NO aplica en
/// <see cref="Buscar(IReadOnlyList{Entrada}, string, decimal, decimal)"/> (no hay verde silencioso
/// ni divergencia fantasma), pero se CONSERVA en <see cref="Entradas"/> como historia auditable
/// (el mecanismo —matcher, reporte, huella— queda intacto; solo cambia el dato). El retiro exige
/// fecha + motivo + evidencia: <see cref="Entrada.RetiroDocumentado"/> lo verifica.
/// </para>
///
/// D-H (Plan 34/40/41/42/46): el matcher NO recibe período y NO ramifica por empresa; recibe la
/// celda y los montos por parámetro.
/// </summary>
public static class DivergenciasDeclaradasCatalogo
{
    /// <summary>Tolerancia de la comparación del monto con el Δ citado (misma doctrina ±0.5).</summary>
    public const decimal Tolerancia = 0.5m;

    /// <summary>
    /// Entrada del catálogo: una divergencia declarada con sus 6 campos citados más el Δ
    /// (diferencia material) que permite reconocerla en la clasificación, y su estado de retiro
    /// (Plan 46 H2: fecha + motivo + evidencia citados; el mecanismo se conserva).
    /// </summary>
    public sealed record Entrada(
        string Id,
        IReadOnlyList<string> Celdas,
        decimal Delta,
        string OrigenMonto,
        string Dueno,
        string CitaEvidencia,
        string Fecha,
        string CondicionRetiro,
        bool Retirada = false,
        string? FechaRetiro = null,
        string? MotivoRetiro = null,
        string? EvidenciaRetiro = null)
    {
        /// <summary>
        /// Verdad si la entrada trae sus 6 campos citados (celda(s), origen del monto, dueño,
        /// cita de evidencia, fecha, condición de retiro) más un Δ no nulo. Una entrada
        /// incompleta NUNCA aplica (R-CAT-2: sin los 6 campos no hay entrada).
        /// </summary>
        public bool Completa =>
            !string.IsNullOrWhiteSpace(Id)
            && Celdas.Count > 0
            && Celdas.All(c => !string.IsNullOrWhiteSpace(c))
            && Delta != 0m
            && !string.IsNullOrWhiteSpace(OrigenMonto)
            && !string.IsNullOrWhiteSpace(Dueno)
            && !string.IsNullOrWhiteSpace(CitaEvidencia)
            && !string.IsNullOrWhiteSpace(Fecha)
            && !string.IsNullOrWhiteSpace(CondicionRetiro);

        /// <summary>
        /// Verdad si el estado de retiro está DOCUMENTADO (fecha + motivo + evidencia citados).
        /// Una entrada vigente (<see cref="Retirada"/> = <c>false</c>) siempre lo está; una retirada
        /// sin documentar es un dato incompleto que el fixture del catálogo rechaza.
        /// </summary>
        public bool RetiroDocumentado =>
            !Retirada
            || (!string.IsNullOrWhiteSpace(FechaRetiro)
                && !string.IsNullOrWhiteSpace(MotivoRetiro)
                && !string.IsNullOrWhiteSpace(EvidenciaRetiro));

        /// <summary>
        /// Verdad si la entrada APLICA hoy: completa, no retirada y con el retiro (si lo hay)
        /// documentado. El matcher solo considera entradas vigentes.
        /// </summary>
        public bool Vigente => Completa && !Retirada && RetiroDocumentado;
    }

    /// <summary>
    /// Catálogo vigente (abierto). Historia completa: la entrada OCCIDENTE-ASE5 queda **RETIRADA**
    /// (Plan 46 H2, decisión del Ingeniero 2026-10-17) porque la divergencia ya no existe. Se
    /// conserva como historia auditable; el matcher la ignora por <see cref="Entrada.Vigente"/>.
    /// </summary>
    public static IReadOnlyList<Entrada> Entradas { get; } =
    [
        new(
            Id: "ASE5-F674-MES2-AUSENTE",
            Celdas: ["F674", "F680"],
            // Δ = 0 POR CONSTRUCCIÓN (Plan 54, D6): la divergencia es ESTRUCTURAL, no material. La
            // fuente de ASE5 trae 2 filas Mes/Total y la geometría v2 declara 3 términos; el slot que
            // la fuente no llena queda en el cero del template y el visible F680 cierra al peso contra
            // el R10 oficial. Con Δ = 0 la entrada NO es `Completa`/`Vigente`, así que documenta y NO
            // matchea: es lo correcto (no hay discrepancia material que reconocer).
            Delta: 0m,
            OrigenMonto:
                "Forma-fuente de 2 filas Mes/Total (ASE5) frente a la geometría v2 de 3 términos: el " +
                "slot Mes2 (F674) no tiene fila de origen y queda en el 0 del template; F680 = " +
                "F674 + F648 + F613 - L613 - L648 = 7.544.253.279,94.",
            Dueno: "UAESP / propietaria del proceso (cita en el T0 del Plan 54)",
            CitaEvidencia:
                "Docs/T0-Plan54-Contrato-Unico-Lectura-R1.md §2 (censo octubre-1: ASE5 filas 23/38, " +
                "cero Aplicacion) y §3 (D9 ASE5 = 7.544.253.279,94; D104 = 7.550.668.664,99 = R10 " +
                "oficial de octubre-1). Fixture: Remuneracion.IntegrationTests/ContratoUnicoLecturaR1T0Tests.cs.",
            Fecha: "2026-10-22",
            CondicionRetiro:
                "El mapa declare slot Mes2 para ASE5 o la forma-fuente traiga 3 filas Mes de modo " +
                "estable; entonces se retira la FILA del catálogo, nunca se adapta el código."),
        new(
            Id: "OCCIDENTE-ASE5",
            Celdas: ["O25", "D34", "F34"],
            Delta: 417519.62m,
            OrigenMonto:
                "Reporte Componentes R1, ancla de OCCIDENTE ASE5: el subtotal del espejo del oficial " +
                "(3 términos; incluye la fila de la 1.ª sección) superaba al del manual legacy " +
                "(2 términos; omite esa fila) en el valor de la 1.ª sección.",
            Dueno: "UAESP",
            CitaEvidencia:
                "Diagnóstico congelado Docs/T0-Plan41-MiniT0-OccidenteAse5-SaldosNota-TablaJ.md §1 " +
                "(ambas fórmulas, valores por término y Δ) y su fixture read-only " +
                "Remuneracion.IntegrationTests/T0FrenteEspejoR1Ase5EnelTests.cs.",
            Fecha: "2026-10-14",
            CondicionRetiro:
                "La UAESP confirme si la remuneración de OCCIDENTE ASE5 incluye la 1.ª sección; " +
                "entonces se elimina esta fila, no se adapta el código.",
            Retirada: true,
            FechaRetiro: "2026-10-17",
            MotivoRetiro:
                "La divergencia YA NO EXISTE: el Plan 46 (sección por frontera de CIERRE + canal " +
                "efectivo) coloca la 1.ª sección de ASE5 en la Aplic#1 oficial, F612 pasa a 0, F700 " +
                "pasa de 305.407.140,43 a 304.989.620,81 (igual a la medición del administrativo) y " +
                "VALIDACION_TOTAL O25/D34/F34 sellan TRUE (Δ 0).",
            EvidenciaRetiro:
                "Plan 46 H2 (plans/46): corrida 9 por CLI (julio-2 desde el oficial) = RESULTADO OK " +
                "codigo=0 sin reporte de divergencia; Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md §12 " +
                "(errata); fixture read-only ContratoSeccionPeriodFreeT0Plan46Tests; " +
                "SeccionEspejoR1Tests.T1_OccidenteAse5_SeccionPorCierre_CoincideConManual.")
    ];

    /// <summary>Entradas que APLICAN hoy (completas, no retiradas); es lo que el matcher consulta.</summary>
    public static IReadOnlyList<Entrada> Vigentes => Entradas.Where(e => e.Vigente).ToList();

    /// <summary>Entradas retiradas (historia auditable; nunca aplican al matcher).</summary>
    public static IReadOnlyList<Entrada> Retiradas => Entradas.Where(e => e.Retirada).ToList();

    /// <summary>
    /// Matcher puro sobre el catálogo vigente. Devuelve la entrada cuya lista de celdas contiene
    /// <paramref name="celda"/> y cuyo Δ coincide con |observado − esperado| (±<see cref="Tolerancia"/>);
    /// <c>null</c> si ninguna entrada VIGENTE aplica (divergencia no declarada o entrada retirada).
    /// </summary>
    public static Entrada? Buscar(string celda, decimal observado, decimal esperado) =>
        Buscar(Entradas, celda, observado, esperado);

    /// <summary>
    /// Matcher puro sobre un catálogo explícito (testeable). Las entradas incompletas o RETIRADAS se
    /// IGNORAN (R-CAT-2: nunca un verde silencioso; Plan 46 H2: una retirada no aplica).
    /// </summary>
    public static Entrada? Buscar(IReadOnlyList<Entrada> catalogo, string celda, decimal observado, decimal esperado)
    {
        if (catalogo is null || string.IsNullOrWhiteSpace(celda))
        {
            return null;
        }

        var delta = Math.Abs(observado - esperado);
        foreach (var entrada in catalogo)
        {
            if (!entrada.Vigente)
            {
                continue;
            }

            if (!entrada.Celdas.Any(c => string.Equals(c, celda, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            if (Math.Abs(delta - entrada.Delta) <= Tolerancia)
            {
                return entrada;
            }
        }

        return null;
    }
}
