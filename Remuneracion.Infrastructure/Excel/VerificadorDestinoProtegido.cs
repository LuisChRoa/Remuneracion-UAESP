using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 59 (H2.2, §2-D2.3 y enmienda §2-D2.3-bis): verificador de DESTINO PROTEGIDO por ROL + FORMA.
///
/// Sustituye las llamadas literales a <c>ValidarCeldaTieneFormula(celda, [fragmentos…])</c>: ya no se
/// compara una coordenada congelada ni una lista de operandos congelada, porque <b>los fragmentos
/// también son geometría</b> (enmienda §2-D2.3-bis: el cierre pasó de <c>F75 = F47+F70+F16-L16-L47</c>
/// en la v2 a <c>F89 = F55+F84+F16-L16-L55</c> en la v3 — cambiaron la celda Y los operandos).
///
/// Cómo se verifica ahora, sin un solo literal de fila/celda:
/// <list type="bullet">
/// <item><b>ROL</b>: el destino se resuelve por identidad conceptual con
/// <see cref="ResolvedorDestinoValor"/> (bloque/ASE + nivel + etiqueta + ocurrencia + encabezado de
/// columna, o tabla + ASE + encabezado). La celda hallada ES la del rol, o el resolvedor falla antes
/// con la identidad completa.</item>
/// <item><b>FORMA</b>: para cada rol se DERIVA el conjunto de términos que su fórmula debe tener a
/// partir de la geometría resuelta: número de términos, signos y destinos salen de los cierres que la
/// plantilla declara (p. ej. el total oportuno del bloque suma los totales de concepto de sus cierres
/// <c>Mes</c> y resta los especiales de todos menos el último). Se compara el multiconjunto
/// <c>(signo, referencia)</c> — nunca el texto literal —, así que una plantilla con más o menos
/// subsecciones se verifica sin cambiar código.</item>
/// </list>
///
/// El guard anti-fórmula del camino de escritura (<see cref="GuardEscrituraFormula"/>) NO se toca: es
/// otra puerta y sigue cruzando todo camino de escritura.
///
/// Mensajes (<c>ERR-PLANTILLA</c> / <c>ERR-SIN-SLOT</c>) por IDENTIDAD, con la celda hallada solo como
/// contexto: un error que se explica con un número de celda es un defecto.
/// </summary>
internal static class VerificadorDestinoProtegido
{
    /// <summary>Título de la tabla AJUSTES - SF-T que alimenta la tabla R5 del CONSOLIDADO.</summary>
    internal const string TituloAjustesSfT = "CONSOLIDADO SALDOS A FAVOR";

    private const string Oportuno = "OPORTUNO";
    private const string Extemporaneo = "EXTEMPORANEO";

    /// <summary>
    /// Verifica las protecciones de las TRES hojas de recaudo en los cinco bloques: total oportuno y
    /// extemporáneo del bloque (R1), total del bloque de anticipos (R2) y total de reversión (R4).
    /// </summary>
    internal static void ValidarBloquesDeRecaudo(WorkbookPart parte)
    {
        ArgumentNullException.ThrowIfNull(parte);

        var r1 = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1);
        var hojaR1 = HojaFisica(parte, DelimitadorBloquesHoja.NombreHojaR1);
        foreach (var bloque in r1)
        {
            ExigirTotalDeBloqueR1(
                parte, hojaR1, bloque, RolTotalBloque.Oportuno,
                CierresDe(bloque, ResolvedorDestinoValor.SubseccionOportuna));
            ExigirTotalDeBloqueR1(
                parte, hojaR1, bloque, RolTotalBloque.Extemporaneo,
                CierresDe(bloque, ResolvedorDestinoValor.SubseccionExtemporanea));
        }

        var r2 = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R2);
        var hojaR2 = HojaFisica(parte, DelimitadorBloquesHoja.NombreHojaR2);
        foreach (var bloque in r2)
        {
            ExigirTotalDeBloqueConSecciones(
                parte, hojaR2, bloque,
                ResolvedorDestinoValor.EncabezadoEspecialesR2,
                "total del bloque de anticipos (R2)");
        }

        var r4 = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R4);
        var hojaR4 = HojaFisica(parte, DelimitadorBloquesHoja.NombreHojaR4);
        foreach (var bloque in r4)
        {
            ExigirTotalDeBloqueR4(parte, hojaR4, bloque);
        }
    }

    /// <summary>
    /// Verifica las 30 celdas del <c>CONSOLIDADO_TOTAL RECAUDO</c> (6 tablas × 5 ASE) y su fila de
    /// cierre: cada una debe referenciar EXACTAMENTE el destino que su tabla declara para ese ASE,
    /// resuelto por identidad — no la coordenada que hoy ocupe ni la que ocupaba ayer.
    /// </summary>
    internal static void ValidarConsolidado(WorkbookPart parte)
    {
        ArgumentNullException.ThrowIfNull(parte);

        var worksheet = HojaFisica(parte, DelimitadorBloquesHoja.NombreConsolidado);
        var r1 = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1);
        var r2 = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R2);
        var r4 = DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R4);
        var hojaR1 = HojaFisica(parte, DelimitadorBloquesHoja.NombreHojaR1);
        var hojaR2 = HojaFisica(parte, DelimitadorBloquesHoja.NombreHojaR2);
        var hojaR4 = HojaFisica(parte, DelimitadorBloquesHoja.NombreHojaR4);
        var hojaAjustes = HojaFisica(parte, DelimitadorBloquesHoja.NombreAjustesSfT);

        foreach (var aseId in IdsDelCatalogo())
        {
            ExigirReferenciaUnica(parte, worksheet, Tablas.OportunoR1, aseId, ResolvedorDestinoValor.ResolverTotalDeBloque(
                parte, hojaR1, DelimitadorBloquesHoja.R1, r1, aseId,
                ResolvedorDestinoValor.RolTotalBloque.Oportuno, ResolvedorDestinoValor.EncabezadoTotal), DelimitadorBloquesHoja.NombreHojaR1);

            ExigirReferenciaUnica(parte, worksheet, Tablas.AnticiposR2, aseId, ResolvedorDestinoValor.ResolverTotalDeBloque(
                parte, hojaR2, DelimitadorBloquesHoja.R2, r2, aseId,
                ResolvedorDestinoValor.RolTotalBloque.Oportuno, ResolvedorDestinoValor.EncabezadoTotal), DelimitadorBloquesHoja.NombreHojaR2);

            ExigirReferenciaUnica(parte, worksheet, Tablas.ExtemporaneoR1, aseId, ResolvedorDestinoValor.ResolverTotalDeBloque(
                parte, hojaR1, DelimitadorBloquesHoja.R1, r1, aseId,
                ResolvedorDestinoValor.RolTotalBloque.Extemporaneo, ResolvedorDestinoValor.EncabezadoTotal), DelimitadorBloquesHoja.NombreHojaR1);

            ExigirReferenciaUnica(parte, worksheet, Tablas.ReversionR4, aseId, ResolvedorDestinoValor.ResolverTotalDeBloque(
                parte, hojaR4, DelimitadorBloquesHoja.R4, r4, aseId,
                ResolvedorDestinoValor.RolTotalBloque.Oportuno, ResolvedorDestinoValor.EncabezadoTotal), DelimitadorBloquesHoja.NombreHojaR4);

            ExigirReferenciaUnica(parte, worksheet, Tablas.SaldosR5, aseId, ResolvedorDestinoValor.ResolverTabla(
                parte, hojaAjustes,
                new ResolvedorDestinoValor.IdentidadTabla(
                    DelimitadorBloquesHoja.NombreAjustesSfT, TituloAjustesSfT, aseId,
                    ResolvedorDestinoValor.EncabezadoRecaudoTotal)), DelimitadorBloquesHoja.NombreAjustesSfT);
        }

        ValidarTablaTotalConsolidado(parte, worksheet);
    }

    /// <summary>
    /// Verifica que los destinos de escritura del <c>DetRetri</c> del período sean celdas de VALOR
    /// (sin <c>&lt;f&gt;</c>): si la plantilla trae fórmula ahí, el escritor values-only no puede pegar
    /// y la corrida aborta nombrando la identidad.
    /// </summary>
    internal static void ValidarDetRetriEditable(WorkbookPart parte, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(periodo);

        var hoja = NombresHojaPeriodo.DetRetri(periodo);
        var worksheet = HojaFisica(parte, hoja);

        foreach (var aseId in IdsDelCatalogo())
        {
            var destino = ResolvedorDestinoValor.ResolverTabla(
                parte, worksheet,
                new ResolvedorDestinoValor.IdentidadTabla(
                    hoja, Titulo: null, aseId, ResolvedorDestinoValor.EncabezadoRecaudoTotal));

            if (destino.CeldaXml.CellFormula is not null)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"ERR-PLANTILLA: el destino de escritura de '{hoja}' para el ASE {aseId} · encabezado '{ResolvedorDestinoValor.EncabezadoRecaudoTotal}' es una fórmula (celda hallada {destino.Celda}). " +
                    "El escritor values-only no puede pegar un valor ahí y la plantilla no debe protegerlo (Plan 59 §2-D2.3).");
            }
        }
    }

    private enum RolTotalBloque
    {
        Oportuno,
        Extemporaneo,
    }

    private static void ExigirTotalDeBloqueR1(
        WorkbookPart parte,
        Worksheet worksheet,
        DelimitadorBloquesHoja.BloqueHoja bloque,
        RolTotalBloque rol,
        IReadOnlyList<DelimitadorBloquesHoja.FilaBloqueHoja> familia)
    {
        var fila = ResolvedorDestinoValor.ResolverTotalDeBloque(
            parte, worksheet, DelimitadorBloquesHoja.R1,
            DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1),
            bloque.AseId,
            rol == RolTotalBloque.Oportuno
                ? ResolvedorDestinoValor.RolTotalBloque.Oportuno
                : ResolvedorDestinoValor.RolTotalBloque.Extemporaneo,
            ResolvedorDestinoValor.EncabezadoTotal).Fila;

        var columnaTotal = ColumnaPorEncabezado(parte, worksheet, bloque, ResolvedorDestinoValor.EncabezadoTotal);
        var columnaEspeciales = ColumnaPorEncabezado(parte, worksheet, bloque, ResolvedorDestinoValor.EncabezadoEspecialesR1);

        var esperados = new List<(int Signo, string Referencia)>();
        for (var i = 0; i < familia.Count; i++)
        {
            esperados.Add((1, $"{columnaTotal}{familia[i].Fila.ToString(CultureInfo.InvariantCulture)}"));
            if (i < familia.Count - 1)
            {
                esperados.Add((-1, $"{columnaEspeciales}{familia[i].Fila.ToString(CultureInfo.InvariantCulture)}"));
            }
        }

        var etiquetaFamilia = familia.Count > 0
            ? familia[0].Etiquetas.B.Trim()
            : (rol == RolTotalBloque.Oportuno
                ? ResolvedorDestinoValor.SubseccionOportuna
                : ResolvedorDestinoValor.SubseccionExtemporanea);

        ExigirForma(
            parte, worksheet, DelimitadorBloquesHoja.NombreHojaR1,
            $"{columnaTotal}{fila.ToString(CultureInfo.InvariantCulture)}",
            esperados,
            $"bloque ASE{bloque.AseId} · sección '{ResolvedorDestinoValor.SeccionDeFila(bloque, fila) ?? "(sin sección)"}' · cuadro de totales · {(rol == RolTotalBloque.Oportuno ? "oportuno" : "extemporáneo")} (Σ '{etiquetaFamilia}' − especiales de todos menos el último)");
    }

    /// <summary>
    /// Localiza el cuadro de totales que vive debajo del gran total del bloque (T0 §2.5: reutiliza las
    /// firmas de cierre, por eso NO se confunde con datos: se busca por su rótulo propio DENTRO de la
    /// región del bloque, que termina en la cabecera del ASE siguiente).
    /// El grupo del bloque se reconoce por <c>C='TOTAL'</c>; su fila oportuna es la que declara
    /// <c>D='OPORTUNO'/E='Total'</c> y su extemporánea la primera <c>D='EXTEMPORANEO'/E='Total'</c> del
    /// mismo grupo (antes de la siguiente fila con rótulo de grupo propio).
    /// </summary>
    private static (int Oportuno, int Extemporaneo) LocalizarCuadroDeTotalesR1(
        WorkbookPart parte,
        Worksheet worksheet,
        DelimitadorBloquesHoja.BloqueHoja bloque)
    {
        var filas = RegionDeTotales(parte, worksheet, bloque);

        var oportunos = filas
            .Where(f => string.Equals(Texto(parte, f, "C"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase)
                && string.Equals(Texto(parte, f, "D"), Oportuno, StringComparison.OrdinalIgnoreCase)
                && string.Equals(Texto(parte, f, "E"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase))
            .Select(f => (int)f.RowIndex!.Value)
            .ToList();

        if (oportunos.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{DelimitadorBloquesHoja.NombreHojaR1}' declara {oportunos.Count} filas de total oportuno del bloque ASE{bloque.AseId} ('{bloque.Rotulo}') en su región de totales y se exige exactamente una (rótulo C='{DelimitadorBloquesHoja.MarcaTotal}' + D='{Oportuno}' + E='{DelimitadorBloquesHoja.MarcaTotal}'). " +
                "El cuadro de totales se identifica por rótulo; nunca por posición.");
        }

        var filaOportuno = oportunos[0];
        var extemporaneos = filas
            .Where(f => f.RowIndex!.Value > filaOportuno)
            .TakeWhile(f => Texto(parte, f, "C").Length == 0)
            .Where(f => string.Equals(Texto(parte, f, "D"), Extemporaneo, StringComparison.OrdinalIgnoreCase)
                && string.Equals(Texto(parte, f, "E"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase))
            .Select(f => (int)f.RowIndex!.Value)
            .ToList();

        if (extemporaneos.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{DelimitadorBloquesHoja.NombreHojaR1}' declara {extemporaneos.Count} filas de total extemporáneo en el grupo del bloque ASE{bloque.AseId} ('{bloque.Rotulo}') y se exige exactamente una (primera D='{Extemporaneo}' + E='{DelimitadorBloquesHoja.MarcaTotal}' tras el total oportuno del grupo).");
        }

        return (filaOportuno, extemporaneos[0]);
    }

    /// <summary>
    /// Filas de la región de totales del bloque: desde la fila siguiente a su gran total hasta la
    /// cabecera del ASE siguiente (o el final de la hoja). La región es estructura (el bloque y su
    /// cuadro los declara la plantilla), no un intervalo congelado.
    /// </summary>
    private static IReadOnlyList<Row> RegionDeTotales(
        WorkbookPart parte,
        Worksheet worksheet,
        DelimitadorBloquesHoja.BloqueHoja bloque) =>
        worksheet.Elements<SheetData>().FirstOrDefault()?.Elements<Row>()
            .Where(r => r.RowIndex is not null && r.RowIndex!.Value > bloque.GranTotal)
            .OrderBy(r => r.RowIndex!.Value)
            .TakeWhile(r => !EsCabeceraDeBloque(parte, r))
            .ToList()
        ?? [];

    // ── R2 / R4 ─────────────────────────────────────────────────────────────────────────────────

    private static void ExigirTotalDeBloqueConSecciones(
        WorkbookPart parte,
        Worksheet worksheet,
        DelimitadorBloquesHoja.BloqueHoja bloque,
        string encabezadoEspeciales,
        string identidadRol)
    {
        var cierres = ResolvedorDestinoValor.CierresDeSeccion(bloque);
        var columnaTotal = ColumnaPorEncabezado(parte, worksheet, bloque, ResolvedorDestinoValor.EncabezadoTotal);
        var columnaEspeciales = ColumnaPorEncabezado(parte, worksheet, bloque, encabezadoEspeciales);
        var filaMaestro = ResolvedorDestinoValor.ResolverTotalDeBloque(
            parte, worksheet, DelimitadorBloquesHoja.R2,
            DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R2),
            bloque.AseId,
            ResolvedorDestinoValor.RolTotalBloque.Oportuno,
            ResolvedorDestinoValor.EncabezadoTotal).Fila;

        var esperados = new List<(int Signo, string Referencia)>();
        for (var i = 0; i < cierres.Count; i++)
        {
            esperados.Add((1, $"{columnaTotal}{cierres[i].Fila.ToString(CultureInfo.InvariantCulture)}"));
            if (i < cierres.Count - 1)
            {
                esperados.Add((-1, $"{columnaEspeciales}{cierres[i].Fila.ToString(CultureInfo.InvariantCulture)}"));
            }
        }

        ExigirForma(
            parte, worksheet, bloque.Hoja,
            $"{columnaTotal}{filaMaestro.ToString(CultureInfo.InvariantCulture)}",
            esperados,
            $"bloque ASE{bloque.AseId} · {identidadRol} (Σ cierres de sección − especiales de todos menos el último)");
    }

    private static void ExigirTotalDeBloqueR4(
        WorkbookPart parte,
        Worksheet worksheet,
        DelimitadorBloquesHoja.BloqueHoja bloque)
    {
        var granTotal = ResolvedorDestinoValor.GranTotal(bloque);
        var columnaTotal = ColumnaPorEncabezado(parte, worksheet, bloque, ResolvedorDestinoValor.EncabezadoTotal);
        var columnaEspeciales = ColumnaPorEncabezado(parte, worksheet, bloque, ResolvedorDestinoValor.EncabezadoEspecialesR4);
        var filaMaestro = ResolvedorDestinoValor.ResolverTotalDeBloque(
            parte, worksheet, DelimitadorBloquesHoja.R4,
            DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R4),
            bloque.AseId,
            ResolvedorDestinoValor.RolTotalBloque.Oportuno,
            ResolvedorDestinoValor.EncabezadoTotal).Fila;
        var fila = granTotal.Fila.ToString(CultureInfo.InvariantCulture);

        ExigirForma(
            parte, worksheet, bloque.Hoja,
            $"{columnaTotal}{filaMaestro.ToString(CultureInfo.InvariantCulture)}",
            [(1, $"{columnaTotal}{fila}"), (-1, $"{columnaEspeciales}{fila}")],
            $"bloque ASE{bloque.AseId} · total de reversión (R4): gran total menos sus especiales");
    }

    /// <summary>
    /// Localiza la fila maestra del cuadro de totales para las gramáticas con marca en <c>B</c> (R2:
    /// <c>B='TOTAL'/C='OPORTUNO'/D='Total'</c>) o en <c>A</c> (R4:
    /// <c>A='TOTAL'/B='OPORTUNO'/C='Total'</c>). Marcas de la gramática del template, no coordenadas.
    /// </summary>
    private static int LocalizarTotalDeBloqueEnCuadro(
        WorkbookPart parte,
        Worksheet worksheet,
        DelimitadorBloquesHoja.BloqueHoja bloque,
        string identidadRol)
    {
        var candidatos = RegionDeTotales(parte, worksheet, bloque)
            .Where(f => string.Equals(Texto(parte, f, "A"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase)
                || string.Equals(Texto(parte, f, "B"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase))
            .Where(f => string.Equals(Texto(parte, f, "B"), Oportuno, StringComparison.OrdinalIgnoreCase)
                || string.Equals(Texto(parte, f, "C"), Oportuno, StringComparison.OrdinalIgnoreCase))
            .Select(f => (int)f.RowIndex!.Value)
            .ToList();

        if (candidatos.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{bloque.Hoja}' declara {candidatos.Count} filas de {identidadRol} en la región de totales del bloque ASE{bloque.AseId} ('{bloque.Rotulo}') y se exige exactamente una (marca '{DelimitadorBloquesHoja.MarcaTotal}' + '{Oportuno}' + '{DelimitadorBloquesHoja.MarcaTotal}').");
        }

        return candidatos[0];
    }

    // ── CONSOLIDADO: tabla total ────────────────────────────────────────────────────────────────

    private static void ValidarTablaTotalConsolidado(WorkbookPart parte, Worksheet worksheet)
    {
        var ids = IdsDelCatalogo();
        var filas = ids
            .Select(id => DestinoDeTabla(parte, worksheet, Tablas.TotalConsolidado, id).Fila)
            .Order()
            .ToList();

        foreach (var aseId in ids)
        {
            var destinoTotal = DestinoDeTabla(parte, worksheet, Tablas.TotalConsolidado, aseId);
            var esperados = Tablas.TodasLasDeRecaudo
                .Select(tabla => DestinoDeTabla(parte, worksheet, tabla, aseId))
                .Select(d => (Signo: 1, Referencia: $"{d.Columna}{d.Fila.ToString(CultureInfo.InvariantCulture)}"))
                .ToList();

            ExigirForma(
                parte, worksheet, DelimitadorBloquesHoja.NombreConsolidado, destinoTotal.Celda, esperados,
                $"CONSOLIDADO · tabla '{Tablas.TotalConsolidado.Titulo}' · fila del ASE {aseId} · suma de las cinco tablas de recaudo");
        }

        var cierre = DestinoDeTabla(parte, worksheet, Tablas.TotalConsolidado, aseId: null);
        ExigirSumaDeRango(
            parte, worksheet, DelimitadorBloquesHoja.NombreConsolidado, cierre,
            filas.First(), filas.Last(),
            $"CONSOLIDADO · tabla '{Tablas.TotalConsolidado.Titulo}' · fila de cierre · suma de las filas por ASE");
    }

    // ── Catálogo de tablas del CONSOLIDADO (por ROL, no por coordenada) ─────────────────────────

    /// <summary>
    /// Tabla por ASE: su <paramref name="Titulo"/> es el rótulo que la plantilla publica en la columna
    /// <c>F</c> y que la identifica (T0 §3.1). Cambiar el título en la plantilla es cambiar el
    /// contrato, y el resolvedor lo detecta con mensaje accionable.
    /// </summary>
    internal sealed record TablaConsolidado(string Titulo);

    internal static class Tablas
    {
        internal static readonly TablaConsolidado OportunoR1 = new("R1 - RECAUDO OPORTUNO");
        internal static readonly TablaConsolidado AnticiposR2 = new("R2 - ANTICIPOS A USUARIOS");
        internal static readonly TablaConsolidado ExtemporaneoR1 = new("R3 - RECAUDO EXTEMPORANEO");
        internal static readonly TablaConsolidado ReversionR4 = new("R4 - REVERSIÓN DE PAGOS");
        internal static readonly TablaConsolidado SaldosR5 = new("R5 - REMUNERACIÓN SALDOS A FAVOR POR NOTA");
        internal static readonly TablaConsolidado TotalConsolidado = new("REPORTE CONSOLIDADO TOTAL DE RECAUDO");

        internal static readonly TablaConsolidado[] TodasLasDeRecaudo =
            [OportunoR1, AnticiposR2, ExtemporaneoR1, ReversionR4, SaldosR5];
    }

    private static ResolvedorDestinoValor.Destino DestinoDeTabla(
        WorkbookPart parte,
        Worksheet worksheet,
        TablaConsolidado tabla,
        int? aseId) =>
        DestinoDeTablaEn(parte, worksheet, DelimitadorBloquesHoja.NombreConsolidado, tabla, aseId);

    private static ResolvedorDestinoValor.Destino DestinoDeTablaEn(
        WorkbookPart parte,
        Worksheet worksheet,
        string hoja,
        TablaConsolidado tabla,
        int? aseId) =>
        ResolvedorDestinoValor.ResolverTabla(
            parte, worksheet,
            new ResolvedorDestinoValor.IdentidadTabla(
                hoja, tabla.Titulo, aseId, ResolvedorDestinoValor.EncabezadoRecaudoTotal));

    /// <summary>
    /// Verifica que la celda de una rendición por empresa referencie el destino resuelto en OTRA hoja.
    /// La referencia esperada se compone con la hoja de origen (identidad del dato), nunca con la
    /// coordenada que el mapa congelado declaraba.
    /// </summary>
    private static void ExigirReferenciaDeOtraHoja(
        WorkbookPart parte,
        Worksheet worksheet,
        string hoja,
        TablaConsolidado tabla,
        int aseId,
        ResolvedorDestinoValor.Destino objetivo,
        string hojaOrigen)
    {
        var destino = DestinoDeTablaEn(parte, worksheet, hoja, tabla, aseId);
        ExigirForma(
            parte, worksheet, hoja, destino.Celda,
            [(1, $"{hojaOrigen.ToUpperInvariant()}!{objetivo.Celda.ToUpperInvariant()}")],
            $"{hoja} · tabla '{tabla.Titulo}' · fila del ASE {aseId} · reflejo del grupo de la empresa en '{hojaOrigen}'");
    }

    private static void ExigirReferenciaUnica(
        WorkbookPart parte,
        Worksheet worksheet,
        TablaConsolidado tabla,
        int aseId,
        ResolvedorDestinoValor.Destino objetivo,
        string hojaOrigen)
    {
        var destino = DestinoDeTabla(parte, worksheet, tabla, aseId);
        ExigirForma(
            parte, worksheet, DelimitadorBloquesHoja.NombreConsolidado, destino.Celda,
            [(1, $"{hojaOrigen.ToUpperInvariant()}!{objetivo.Celda.ToUpperInvariant()}")],
            $"CONSOLIDADO · tabla '{tabla.Titulo}' · fila del ASE {aseId} · reflejo del destino que declara la tabla origen '{hojaOrigen}'");
    }

    // ── Verificación de FORMA ───────────────────────────────────────────────────────────────────

    private static void ExigirForma(
        WorkbookPart parte,
        Worksheet worksheet,
        string hoja,
        string celda,
        IReadOnlyList<(int Signo, string Referencia)> esperados,
        string identidad)
    {
        var cell = ResolvedorDestinoValor.BuscarCelda(worksheet, celda)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: el destino de {identidad} no existe en la hoja '{hoja}' (celda hallada: {celda}). La plantilla no declara el slot.");

        if (cell.CellFormula is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidad} debería ser una fórmula del template (celda hallada {celda} de '{hoja}') y la plantilla trae un valor fijo. " +
                "La protección se verifica por ROL y FORMA sobre la geometría que la plantilla declara.");
        }

        // Shared-follower: réplica del maestro sin texto propio. La forma ya se verifica en el maestro
        // (misma fórmula desplazada); aquí basta la presencia del rol.
        if (string.IsNullOrWhiteSpace(cell.CellFormula.Text) && cell.CellFormula.SharedIndex is not null)
        {
            return;
        }

        var actuales = Terminos(cell.CellFormula.Text, hoja)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidad} (celda hallada {celda} de '{hoja}') tiene una fórmula que no es una suma/resta de referencias ('{cell.CellFormula.Text}'): no se puede verificar su forma sin comparar texto literal, y este verificador no compara literales.");

        var esperado = esperados
            .Select(t => (t.Signo, Referencia: t.Referencia.Contains('!')
                ? NormalizarReferencia(t.Referencia, hoja)
                : NormalizarFormula(t.Referencia)))
            .OrderBy(t => t.Referencia, StringComparer.Ordinal)
            .ThenBy(t => t.Signo)
            .ToList();
        var obtenido = actuales.OrderBy(t => t.Referencia, StringComparer.Ordinal).ThenBy(t => t.Signo).ToList();

        if (esperado.Count != obtenido.Count || !esperado.SequenceEqual(obtenido))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidad}: la fórmula no tiene la forma que su geometría exige (celda hallada {celda} de '{hoja}'). " +
                $"Esperado por forma: {DescribirTerminos(esperado)}; obtenido: {DescribirTerminos(obtenido)}. " +
                "La forma se deriva de los cierres que la plantilla declara; no se compara una lista de operandos congelada (§2-D2.3-bis).");
        }
    }

    private static void ExigirSumaDeRango(
        WorkbookPart parte,
        Worksheet worksheet,
        string hoja,
        ResolvedorDestinoValor.Destino cierre,
        int filaDesde,
        int filaHasta,
        string identidad)
    {
        _ = parte;
        var cell = cierre.CeldaXml;
        if (cell.CellFormula is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidad} debería ser una fórmula del template (celda hallada {cierre.Celda} de '{hoja}') y la plantilla trae un valor fijo.");
        }

        if (string.IsNullOrWhiteSpace(cell.CellFormula.Text) && cell.CellFormula.SharedIndex is not null)
        {
            return;
        }

        var columna = cierre.Columna;
        var esperado = $"SUM({columna}{filaDesde.ToString(CultureInfo.InvariantCulture)}:{columna}{filaHasta.ToString(CultureInfo.InvariantCulture)})";
        var texto = cell.CellFormula.Text.Replace(" ", string.Empty, StringComparison.Ordinal).ToUpperInvariant();
        if (!string.Equals(texto, esperado.ToUpperInvariant(), StringComparison.Ordinal))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidad}: se esperaba la suma del rango que la tabla declara ('{esperado}') y la fórmula es '{cell.CellFormula.Text}' (celda hallada {cierre.Celda} de '{hoja}').");
        }
    }

    /// <summary>
    /// Descompone una fórmula de suma/resta de referencias en su multiconjunto <c>(signo, referencia)</c>.
    /// Devuelve <c>null</c> cuando la fórmula no es de esa familia (paréntesis, funciones, literales):
    /// el verificador no adivina formas.
    /// </summary>
    private static List<(int Signo, string Referencia)>? Terminos(string? formula, string hojaPorDefecto)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return null;
        }

        var texto = formula.Replace(" ", string.Empty, StringComparison.Ordinal);
        if (texto.IndexOfAny(['(', ')', '*', '/', ';', ',']) >= 0)
        {
            return null;
        }

        var terminos = new List<(int, string)>();
        var actual = new System.Text.StringBuilder();
        var signo = 1;
        var i = 0;
        while (i < texto.Length)
        {
            var caracter = texto[i];

            // Un nombre de hoja entrecomillado puede contener '+' o '-' (p. ej. 'AJUSTES - SF-T'):
            // se copia entero para no partirlo en términos falsos.
            if (caracter == '\'')
            {
                var cierre = texto.IndexOf('\'', i + 1);
                if (cierre < 0)
                {
                    return null;
                }

                actual.Append(texto, i, cierre - i + 1);
                i = cierre + 1;
                continue;
            }

            if (caracter is '+' or '-')
            {
                if (actual.Length == 0)
                {
                    if (i == 0 && caracter == '-')
                    {
                        signo = -1;
                        i++;
                        continue;
                    }

                    return null;
                }

                terminos.Add((signo, NormalizarReferencia(actual.ToString(), hojaPorDefecto)));
                actual.Clear();
                signo = caracter == '-' ? -1 : 1;
                i++;
                continue;
            }

            actual.Append(caracter);
            i++;
        }

        if (actual.Length == 0)
        {
            return null;
        }

        terminos.Add((signo, NormalizarReferencia(actual.ToString(), hojaPorDefecto)));
        return terminos;
    }

    private static string NormalizarReferencia(string termino, string hojaPorDefecto)
    {
        var separador = termino.LastIndexOf('!');
        if (separador < 0)
        {
            return NormalizarFormula(termino);
        }

        var hoja = NormalizarFormula(termino[..separador].Trim('\''));
        var celda = NormalizarFormula(termino[(separador + 1)..]);
        return string.Equals(hoja, NormalizarFormula(hojaPorDefecto), StringComparison.Ordinal)
            ? celda
            : $"{hoja}!{celda}";
    }

    /// <summary>
    /// Normaliza un fragmento de fórmula para comparar FORMA: sin espacios, sin comillas de hoja y con
    /// las variantes tipográficas del guion unificadas. El guion importa porque el texto guardado por
    /// Excel puede usar <c>U+2212</c> para el mismo nombre de hoja que la plantilla declara con
    /// <c>-</c>; unificar el carácter es normalización de texto, no una concesión de coordenada.
    /// </summary>
    private static string NormalizarFormula(string valor) =>
        valor.Replace(" ", string.Empty, StringComparison.Ordinal)
            .Replace("−", "-", StringComparison.Ordinal)
            .Replace("–", "-", StringComparison.Ordinal)
            .Replace("—", "-", StringComparison.Ordinal)
            .ToUpperInvariant();

    private static string DescribirTerminos(IReadOnlyList<(int Signo, string Referencia)> terminos) =>
        terminos.Count == 0
            ? "(ninguno)"
            : string.Join(" ", terminos.Select(t => $"{(t.Signo > 0 ? "+" : "−")}{t.Referencia}"));

    // ── Utilidades ──────────────────────────────────────────────────────────────────────────────

    private static IReadOnlyList<DelimitadorBloquesHoja.FilaBloqueHoja> CierresDe(
        DelimitadorBloquesHoja.BloqueHoja bloque,
        string etiqueta) =>
        ResolvedorDestinoValor.CierresDeSubseccion(bloque, etiqueta);

    private static string PrimerCierreDeSeccion(
        IReadOnlyList<DelimitadorBloquesHoja.BloqueHoja> bloques,
        int aseId)
    {
        var bloque = bloques.FirstOrDefault(b => b.AseId == aseId)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{DelimitadorBloquesHoja.NombreHojaR2}' no declara el bloque del ASE {aseId}.");
        return ResolvedorDestinoValor.CierresDeSeccion(bloque).FirstOrDefault()?.Etiquetas.A.Trim()
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: el bloque ASE{aseId} de '{bloque.Hoja}' no declara ningún cierre de sección.");
    }

    private static string ColumnaPorEncabezado(
        WorkbookPart parte,
        Worksheet worksheet,
        DelimitadorBloquesHoja.BloqueHoja bloque,
        string encabezado)
    {
        var fila = worksheet.Descendants<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == bloque.EncabezadoColumnas)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: el bloque ASE{bloque.AseId} de '{bloque.Hoja}' no declara fila de encabezados de columna: no se puede resolver la columna '{encabezado}'.");

        foreach (var celda in fila.Elements<Cell>())
        {
            var letra = ResolvedorDestinoValor.LetraDeColumna(celda.CellReference?.Value);
            if (letra.Length > 0
                && string.Equals(LectorTextoHoja.Texto(parte, celda).Trim(), encabezado, StringComparison.OrdinalIgnoreCase))
            {
                return letra;
            }
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-SIN-SLOT: el bloque ASE{bloque.AseId} de '{bloque.Hoja}' no declara la columna '{encabezado}' que exige la verificación de destinos protegidos.");
    }

    private static string Texto(WorkbookPart parte, Row fila, string columna) =>
        LectorTextoHoja.TextoDeColumna(parte, fila, columna).Trim();

    private static bool EsCabeceraDeBloque(WorkbookPart parte, Row fila)
    {
        var texto = Texto(parte, fila, "B");
        if (texto.Length == 0)
        {
            return false;
        }

        for (var id = 1; id <= CarpetasAse.Prefijos.Count; id++)
        {
            if (SinonimosEmpresaR1.SonMismaEmpresa(texto, AseFactory.DesdeId(id).NombreCompleto))
            {
                return true;
            }
        }

        return false;
    }

    private static Worksheet HojaFisica(WorkbookPart parte, string nombre)
    {
        ArgumentNullException.ThrowIfNull(parte);
        var hoja = parte.Workbook?.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: el workbook no trae la hoja '{nombre}' que exige la verificación de destinos protegidos.");
        return ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{nombre}' no tiene WorksheetPart válido.");
    }

    private static IReadOnlyList<int> IdsDelCatalogo() =>
        Enumerable.Range(1, CarpetasAse.Prefijos.Count).ToList();
}
