namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 21 (T5): celdas L del mapa legado rol/ocurrencia (T0-0.5 de HU-16) congeladas como
/// EVIDENCIA de test. El mapa productivo (<c>WorkbookLeafCellMapInterventoria.LMenoresPorAse</c>/
/// <c>LMenoresPorAseQ2</c>), el enum <c>RolLMenor</c>, <c>ObtenerLMenores</c> y
/// <c>LeerLEspecialesMenores</c> se retiraron al probarse la absorción 15/15 por el espejo
/// estructural R1. Estas listas permiten que los tests de Capa A y de absorción sigan verificando
/// las mismas celdas contra golden/plantilla sin referenciar tipos eliminados.
/// </summary>
internal static class MapaLMenoresEvidenciaT0
{
    internal static readonly IReadOnlyDictionary<int, string[]> Q1 = new Dictionary<int, string[]>
    {
        [1] = ["L11", "L12", "L13", "L21", "L22", "L23", "L26", "L43"],
        [2] = ["L92", "L93", "L94", "L106", "L107", "L108", "L109", "L111", "L114", "L132"],
        [3] = ["L224", "L225", "L226", "L231", "L233", "L234", "L235", "L236", "L239", "L256"],
        [4] = ["L370", "L371", "L372", "L384", "L385", "L386", "L387", "L389", "L392", "L419"],
        [5] = ["L480", "L481", "L482", "L485", "L486", "L494", "L495", "L496", "L499", "L515"]
    };

    internal static readonly IReadOnlyDictionary<int, string[]> Q2 = new Dictionary<int, string[]>
    {
        [1] = ["L18", "L19", "L20", "L21", "L25", "L27", "L28", "L29", "L30", "L31", "L33", "L50"],
        [2] = ["L108", "L109", "L110", "L111", "L122", "L123", "L124", "L125", "L127", "L128", "L129", "L130", "L131", "L133", "L162"],
        [3] = ["L251", "L252", "L253", "L254", "L255", "L257", "L258", "L260", "L261", "L262", "L263", "L264", "L266", "L283"],
        [4] = ["L398", "L399", "L400", "L401", "L415", "L417", "L418", "L419", "L420", "L421", "L423", "L450"],
        [5] = ["L513", "L514", "L515", "L516", "L524", "L526", "L527", "L528", "L529", "L530", "L532", "L554"]
    };

    internal static IReadOnlyDictionary<int, string[]> PorQuincena(int numeroQuincena) =>
        numeroQuincena == 2 ? Q2 : Q1;
}
