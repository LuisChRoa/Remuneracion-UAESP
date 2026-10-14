using Remuneracion.Infrastructure.Excel;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 38 (T3, R-B-4 / D-C): cita de las direcciones LEGACY de las celdas-operando R1/R2/R4-Q2
/// que viven en el caché golden (<c>Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total
/// Administrativo.xlsx</c> = manual del administrativo, plantilla LEGACY).
///
/// El runtime escribe sobre la plantilla estandarizada ÚNICA (geometría nueva, mini-T0
/// <see cref="GeometriaPlantillaEstandarizadaT0Tests"/>); el golden es la referencia de REGRESIÓN
/// del manual y conserva la geometría legacy. Para que los tests estructurales sigan comparando
/// VALORES (su contrato) sin ajustar números, se empareja por ROL: el golden se lee en la celda
/// legacy y el leaf en la celda nueva del mapa.
///
/// Evidencia de los roles (dump BCL de la fórmula visible legacy, 2026-10-14):
/// - R1-Q2: <c>F53=F32+F48+F12-L12-L32</c> (ASE1) y <c>F55=F37+F17-L17</c>.
/// - R2-Q2: <c>E43=E17+E28-K17</c> (ASE1), <c>E438=E410+E396-K396</c> (ASE5).
/// - R4-Q2: <c>D73=D15-P15</c>, <c>D168=D105-P105</c>, …, <c>D355=D352-P352</c>.
///
/// NO es producción: solo cita de comparación contra el golden legacy (declarado).
/// </summary>
internal static class GeometriaGoldenLegacy
{
    /// <summary>Celda legacy del caché golden para cada rol R1-Q2 por ASE.</summary>
    internal static readonly Dictionary<(int Ase, WorkbookLeafCellMapQ2.R1Q2Fuente Rol), string> R1Q2 =
        new()
        {
            [(1, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes0)] = "F12",
            [(1, WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes0)] = "L12",
            [(1, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes1)] = "F32",
            [(1, WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes1)] = "L32",
            [(1, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes2)] = "F48",
            [(1, WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic0)] = "F17",
            [(1, WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic1)] = "F37",
            [(1, WorkbookLeafCellMapQ2.R1Q2Fuente.LAplic0)] = "L17",
            [(2, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes0)] = "F94",
            [(2, WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes0)] = "L94",
            [(2, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes1)] = "F132",
            [(2, WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes1)] = "L132",
            [(2, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes2)] = "F160",
            [(2, WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic0)] = "F107",
            [(2, WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic1)] = "F142",
            [(2, WorkbookLeafCellMapQ2.R1Q2Fuente.LAplic0)] = "L107",
            [(3, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes0)] = "F244",
            [(3, WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes0)] = "L244",
            [(3, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes1)] = "F265",
            [(3, WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes1)] = "L265",
            [(3, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes2)] = "F281",
            [(3, WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic0)] = "F250",
            [(3, WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic1)] = "F270",
            [(3, WorkbookLeafCellMapQ2.R1Q2Fuente.LAplic0)] = "L250",
            [(4, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes0)] = "F384",
            [(4, WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes0)] = "L384",
            [(4, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes1)] = "F422",
            [(4, WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes1)] = "L422",
            [(4, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes2)] = "F448",
            [(4, WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic0)] = "F397",
            [(4, WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic1)] = "F430",
            [(4, WorkbookLeafCellMapQ2.R1Q2Fuente.LAplic0)] = "L397",
            [(5, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes0)] = "F531",
            [(5, WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes0)] = "L531",
            [(5, WorkbookLeafCellMapQ2.R1Q2Fuente.Mes1)] = "F552",
            [(5, WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic0)] = "F512",
            [(5, WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic1)] = "F538",
            [(5, WorkbookLeafCellMapQ2.R1Q2Fuente.LAplic0)] = "L512"
        };

    /// <summary>Celda legacy R2-Q2 (Componente / SubsCont / Especiales) por ASE en el golden.</summary>
    internal static readonly Dictionary<int, (string Componente, string SubsCont, string Especiales)> R2Q2 =
        new()
        {
            [1] = ("E17", "E28", "K17"),
            [2] = ("E89", "E107", "K89"),
            [3] = ("E176", "E187", "K176"),
            [4] = ("E303", "E323", "K303"),
            // ASE5: la fórmula visible legacy E438 = E410+E396-K396 trae el SubsCont ANTES del
            // Componente (E410=SubsCont=4810642.25, E396=Componente=39162866.71). Evidencia del
            // dump del golden; el rol se fija por el rótulo de la fila, no por el orden del operando.
            [5] = ("E396", "E410", "K396")
        };

    /// <summary>Celda legacy R4-Q2 (Total / P) por ASE en el golden.</summary>
    internal static readonly Dictionary<int, (string Total, string P)> R4Q2 =
        new()
        {
            [1] = ("D15", "P15"),
            [2] = ("D105", "P105"),
            [3] = ("D200", "P200"),
            [4] = ("D243", "P243"),
            [5] = ("D352", "P352")
        };
}
