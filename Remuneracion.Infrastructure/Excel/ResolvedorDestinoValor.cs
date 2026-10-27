using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 59 (H2.1, §2-D2): RESOLUTOR DE DESTINO DE VALOR por IDENTIDAD CONCEPTUAL.
///
/// Reemplaza la localización por coordenada congelada (los mapas <c>WorkbookLeafCellMapQ2.*</c> y
/// <c>WorkbookLeafCellMapPorEmpresa.ProtectedFormulasPorEmpresa</c>, retirados del camino productivo
/// en este H): un destino se pide por lo que ES, no por dónde está.
/// <list type="bullet">
/// <item><b>Nivel de bloque</b> (hojas de recaudo R1/R2/R4): <c>(hoja, ASE, nivel de cierre, sección
/// contenedora, etiqueta del cierre, ocurrencia, encabezado de columna)</c>. La geometría la aporta
/// <see cref="DelimitadorBloquesHoja"/> (H1), que la lee de la propia plantilla; la columna se
/// resuelve por el ENCABEZADO declarado en la fila de encabezados del bloque (Plan 44), nunca por
/// índice ni por letra fija.</item>
/// <item><b>Nivel de tabla</b> (CONSOLIDADO, AJUSTES - SF-T, DetRetri): <c>(hoja, título de la tabla,
/// ASE, encabezado de columna)</c> con la regla medida en el T0 (§3.1): el título vive en <c>F</c> y
/// la tabla se reconoce porque DEBAJO declara un encabezado <c>B='ASE'</c>; la fila se identifica por
/// el id numérico del ASE leído del propio workbook (H-59-4: el nombre largo puede diferir entre
/// hojas por un punto, así que NO se compara contra el de otra hoja).</item>
/// </list>
///
/// Prohibiciones que este tipo hace estructuralmente imposibles:
/// <list type="bullet">
/// <item>cero literales de fila/celda: ninguna ruta de este archivo construye una referencia con un
/// número escrito, solo con la fila que el delimitador leyó y la columna que el encabezado nombró;</item>
/// <item>cero ramas por ASE/empresa/período/quincena: el ASE entra como id de dominio y se contrasta
/// contra el rótulo de la cabecera de bloque (<see cref="DelimitadorBloquesHoja.DelimitarBloque"/>);</item>
/// <item>cero offsets: una ocurrencia (<c>k</c>-ésima etiqueta dentro de su contenedor) es identidad,
/// no aritmética de filas.</item>
/// </list>
///
/// Fail-fast accionable (<c>ERR-SIN-SLOT</c> / <c>ERR-PLANTILLA</c>) con la IDENTIDAD completa
/// (bloque · sección · subsección · ocurrencia · encabezado) y la celda hallada solo como contexto:
/// un error que se explica con un número de celda es un defecto (§2-D2.3-bis).
/// </summary>
internal static class ResolvedorDestinoValor
{
    /// <summary>Nivel de cierre que identifica el destino dentro de un bloque de recaudo.</summary>
    internal enum NivelDestino
    {
        /// <summary>Cierre de sección (<c>A=&lt;sección&gt;/B='Total'</c>).</summary>
        Seccion,

        /// <summary>Cierre de subsección/cuadro (<c>B=&lt;subsección&gt;/C='Total'</c>). Solo R1.</summary>
        Subseccion,

        /// <summary>Gran total del bloque (<c>A='Total'</c> con <c>B</c> vacía).</summary>
        GranTotal,
    }

    /// <summary>
    /// Identidad conceptual de un destino dentro de una hoja de recaudo.
    /// <paramref name="Ocurrencia"/> es la k-ésima aparición de <paramref name="Etiqueta"/> dentro
    /// del BLOQUE (1-based) y <paramref name="Encabezado"/> es el rótulo de columna del bloque.
    /// </summary>
    internal sealed record Identidad(
        string Hoja,
        int AseId,
        NivelDestino Nivel,
        string? Etiqueta,
        int Ocurrencia,
        string Encabezado)
    {
        /// <summary>Descripción por identidad, para mensajes: nunca empieza por una coordenada.</summary>
        internal string Describir() => Nivel switch
        {
            NivelDestino.Seccion =>
                $"bloque ASE{AseId} · sección '{Etiqueta}' · encabezado '{Encabezado}'",
            NivelDestino.Subseccion =>
                $"bloque ASE{AseId} · subsección '{Etiqueta}' #{Ocurrencia} · encabezado '{Encabezado}'",
            NivelDestino.GranTotal =>
                $"bloque ASE{AseId} · gran total · encabezado '{Encabezado}'",
            _ => $"bloque ASE{AseId} · '{Etiqueta}' · encabezado '{Encabezado}'",
        };
    }

    /// <summary>Destino resuelto: fila, letra de columna, referencia A1 y la celda del template.</summary>
    internal sealed record Destino(int Fila, string Columna, string Celda, Cell CeldaXml);

    /// <summary>
    /// Identidad de un destino dentro de una TABLA por ASE (CONSOLIDADO / AJUSTES - SF-T / DetRetri).
    /// <paramref name="Titulo"/> es el rótulo que la tabla declara en <c>F</c> (null cuando la hoja
    /// declara una única tabla y su encabezado es suficiente); <paramref name="AseId"/> null pide la
    /// fila de cierre de la tabla (<c>C='TOTAL'</c>); <paramref name="Encabezado"/> es el rótulo de la
    /// columna de valor dentro del encabezado de la tabla.
    /// </summary>
    internal sealed record IdentidadTabla(string Hoja, string? Titulo, int? AseId, string Encabezado)
    {
        /// <summary>Descripción por identidad, para mensajes.</summary>
        internal string Describir() => AseId is { } ase
            ? $"tabla '{Titulo ?? Hoja}' de '{Hoja}' · fila del ASE {ase} · encabezado '{Encabezado}'"
            : $"tabla '{Titulo ?? Hoja}' de '{Hoja}' · fila de cierre · encabezado '{Encabezado}'";
    }

    /// <summary>Encabezado de la columna de valor total en las tres hojas de recaudo y en las tablas.</summary>
    internal const string EncabezadoTotal = "Total";

    /// <summary>Encabezado de la columna de servicio especiales en R1.</summary>
    internal const string EncabezadoEspecialesR1 = "SERVICIO ESPECIALES";

    /// <summary>Encabezado de la columna de especiales en R2.</summary>
    internal const string EncabezadoEspecialesR2 = "Especiales";

    /// <summary>Encabezado de la columna de servicio especiales en R4.</summary>
    internal const string EncabezadoEspecialesR4 = "SERVICIO ESPECIALES";

    /// <summary>Encabezado de la columna de valor de las tablas por ASE (CONSOLIDADO/AJUSTES/DetRetri).</summary>
    internal const string EncabezadoRecaudoTotal = "Recaudo Total";    /// <summary>Rótulo del encabezado de tabla por ASE en su columna <c>B</c> (gramática del template).</summary>
    internal const string EncabezadoTablaAse = "ASE";

    /// <summary>Rótulo de la fila de cierre de una tabla por ASE (columna <c>C</c>).</summary>
    internal const string MarcaTotalTabla = "TOTAL";

    /// <summary>Subsección cuyo total de cierre alimenta el TOT_OPT del bloque (recaudo del mes).</summary>
    internal const string SubseccionOportuna = "Mes";

    /// <summary>Subsección cuyo total de cierre alimenta el EXTEMP del bloque (aplicación nuevos x reversión).</summary>
    internal const string SubseccionExtemporanea = "Aplicacion nuevos x reversion";

    // ── Nivel bloque (R1/R2/R4) ──────────────────────────────────────────────────────────────────

    /// <summary>
    /// Resuelve el destino de un cierre dentro del bloque del ASE. La búsqueda recorre las filas que
    /// el delimitador YA clasificó por firma: no hay ningún cálculo de posición.
    /// </summary>
    internal static Destino Resolver(
        WorkbookPart parte,
        Worksheet hojaFisica,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        IReadOnlyList<DelimitadorBloquesHoja.BloqueHoja> bloques,
        Identidad identidad)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(hojaFisica);
        ArgumentNullException.ThrowIfNull(hoja);
        ArgumentNullException.ThrowIfNull(bloques);
        ArgumentNullException.ThrowIfNull(identidad);

        var bloque = bloques.FirstOrDefault(b => b.AseId == identidad.AseId)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{hoja.Nombre}' no declara el bloque del ASE {identidad.AseId} que exige el destino {identidad.Describir()}. Bloques hallados: {string.Join(", ", bloques.Select(b => $"ASE {b.AseId} '{b.Rotulo}'"))}.");

        var fila = LocalizarFila(bloque, hoja, identidad);
        var columna = ColumnaPorEncabezado(parte, hojaFisica, bloque, identidad);
        var referencia = $"{columna}{fila.ToString(CultureInfo.InvariantCulture)}";
        var celda = BuscarCelda(hojaFisica, referencia)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: el destino {identidad.Describir()} se resolvió a la celda {referencia} de la hoja '{hoja.Nombre}', pero esa celda no existe en la plantilla. No se inventa: la plantilla debe declararla.");

        return new Destino(fila, columna, referencia, celda);
    }

    /// <summary>Filas de cierre que la identidad referencia (una o varias ocurrencias del mismo rol).</summary>
    internal static IReadOnlyList<DelimitadorBloquesHoja.FilaBloqueHoja> CierresDeSubseccion(
        DelimitadorBloquesHoja.BloqueHoja bloque,
        string etiqueta)
    {
        ArgumentNullException.ThrowIfNull(bloque);
        return bloque.Filas
            .Where(f => f.Tipo == DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSubseccion
                && string.Equals(f.Etiquetas.B.Trim(), etiqueta.Trim(), StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    /// <summary>Filas de cierre de sección del bloque, en orden físico.</summary>
    internal static IReadOnlyList<DelimitadorBloquesHoja.FilaBloqueHoja> CierresDeSeccion(
        DelimitadorBloquesHoja.BloqueHoja bloque)
    {
        ArgumentNullException.ThrowIfNull(bloque);
        return bloque.Cierres(DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion);
    }

    /// <summary>
    /// Sección a la que pertenece una fila del bloque (la que la CIERRA, primera hacia abajo). Es
    /// contexto para los mensajes: la identidad del destino no depende de la sección, pero el error
    /// debe nombrarla para que la propietaria lo lea en su vocabulario.
    /// </summary>
    internal static string? SeccionDeFila(DelimitadorBloquesHoja.BloqueHoja bloque, int fila)
    {
        ArgumentNullException.ThrowIfNull(bloque);
        return bloque.Filas
            .Where(f => f.Fila >= fila && f.Tipo == DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion)
            .Select(f => f.Etiquetas.A.Trim())
            .FirstOrDefault();
    }

    /// <summary>Fila del gran total del bloque (una sola: la que corta el bloque).</summary>
    internal static DelimitadorBloquesHoja.FilaBloqueHoja GranTotal(DelimitadorBloquesHoja.BloqueHoja bloque)
    {
        ArgumentNullException.ThrowIfNull(bloque);
        return bloque.Filas.First(f => f.Tipo == DelimitadorBloquesHoja.TipoFilaBloqueHoja.GranTotal);
    }

    private static int LocalizarFila(
        DelimitadorBloquesHoja.BloqueHoja bloque,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        Identidad identidad)
    {
        switch (identidad.Nivel)
        {
            case NivelDestino.GranTotal:
                return GranTotal(bloque).Fila;

            case NivelDestino.Seccion:
            {
                var ocurrencia = 0;
                foreach (var fila in CierresDeSeccion(bloque))
                {
                    if (!string.Equals(fila.Etiquetas.A.Trim(), identidad.Etiqueta?.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (++ocurrencia == Math.Max(identidad.Ocurrencia, 1))
                    {
                        return fila.Fila;
                    }
                }

                throw SinDestino(hoja, bloque, identidad);
            }

            case NivelDestino.Subseccion:
            {
                // Ocurrencia a nivel de BLOQUE: el cierre de subsección se identifica por su etiqueta
                // y su orden dentro del bloque (las fórmulas del template suman las N ocurrencias,
                // sin importar a qué sección pertenece cada una).
                var ocurrencia = 0;
                foreach (var fila in bloque.Filas)
                {
                    if (fila.Tipo != DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSubseccion
                        || !string.Equals(fila.Etiquetas.B.Trim(), identidad.Etiqueta?.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (++ocurrencia == Math.Max(identidad.Ocurrencia, 1))
                    {
                        return fila.Fila;
                    }
                }

                throw SinDestino(hoja, bloque, identidad);
            }

            default:
                throw SinDestino(hoja, bloque, identidad);
        }
    }

    /// <summary>
    /// Rol de un valor editable: qué ES el número que se va a pegar, sin decir dónde va. Es la
    /// llave con la que el lector entrega los valores y el escritor los resuelve contra la plantilla
    /// (el transporte entre ambos NO lleva referencias de celda: llevarlas era el mapa congelado).
    /// </summary>
    internal readonly record struct RolValor(NivelDestino Nivel, string Etiqueta, int Ocurrencia, string Columna)
    {
        /// <summary>Clave estable del rol (la que viaja en <c>ValoresPorRol</c>).</summary>
        internal string Clave => $"{Nivel}|{Etiqueta}|{Ocurrencia.ToString(CultureInfo.InvariantCulture)}|{Columna}";

        /// <summary>Identidad de destino del rol en un bloque concreto.</summary>
        internal Identidad IdentidadEn(string hoja, int aseId) => new(hoja, aseId, Nivel, Etiqueta, Ocurrencia, Columna);

        /// <summary>Reconstruye el rol desde su clave; clave mal formada = defecto de programación.</summary>
        internal static RolValor DesdeClave(string clave)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(clave);
            var partes = clave.Split('|');
            if (partes.Length != 4
                || !Enum.TryParse<NivelDestino>(partes[0], ignoreCase: false, out var nivel)
                || !int.TryParse(partes[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out var ocurrencia))
            {
                throw new ArgumentException($"Clave de rol de destino inválida: '{clave}'.", nameof(clave));
            }

            return new RolValor(nivel, partes[1], ocurrencia, partes[3]);
        }
    }

    /// <summary>
    /// Columna destino por ENCABEZADO declarado en la fila de encabezados del bloque. La primera
    /// ocurrencia manda (Plan 44/T0: el rótulo repetido conserva la columna de la izquierda), y un
    /// encabezado ausente es fail-fast: la plantilla no declara ese destino.
    /// </summary>
    private static string ColumnaPorEncabezado(
        WorkbookPart parte,
        Worksheet hojaFisica,
        DelimitadorBloquesHoja.BloqueHoja bloque,
        Identidad identidad)
    {
        var fila = hojaFisica.Descendants<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == bloque.EncabezadoColumnas)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la fila de encabezados {bloque.EncabezadoColumnas} del bloque ASE{bloque.AseId} no existe en la hoja física y el destino {identidad.Describir()} exige resolver la columna por encabezado.");

        foreach (var celda in fila.Elements<Cell>())
        {
            var letra = LetraDeColumna(celda.CellReference?.Value);
            if (letra.Length == 0)
            {
                continue;
            }

            var texto = LectorTextoHoja.Texto(parte, celda).Trim();
            if (string.Equals(texto, identidad.Encabezado.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return letra;
            }
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-SIN-SLOT: el bloque ASE{bloque.AseId} no declara una columna con el encabezado '{identidad.Encabezado}' (destino pedido: {identidad.Describir()}). Encabezados declarados: {string.Join(", ", fila.Elements<Cell>().Select(c => LectorTextoHoja.Texto(parte, c).Trim()).Where(t => t.Length > 0))}.");
    }

    /// <summary>
    /// Rol del cuadro de totales que vive debajo del gran total del bloque (T0 §2.5: reutiliza las
    /// firmas de cierre, por eso se identifica por su rótulo propio y no por vecindad de filas).
    /// </summary>
    internal enum RolTotalBloque
    {
        /// <summary>Total oportuno del bloque (<c>D='OPORTUNO'/E='Total'</c> en el grupo <c>C='TOTAL'</c>).</summary>
        Oportuno,

        /// <summary>Total extemporáneo del bloque (<c>D='EXTEMPORANEO'/E='Total'</c> del mismo grupo).</summary>
        Extemporaneo,
    }

    /// <summary>
    /// Resuelve el destino del TOTAL DEL BLOQUE (la celda que consumen las tablas del CONSOLIDADO y las
    /// rendiciones por empresa). R1 lo declara en su cuadro de totales con <c>C='TOTAL'</c>; R2 con
    /// <c>B='TOTAL'</c>; R4 con <c>A='TOTAL'</c>. La región se acota por el gran total del bloque y la
    /// cabecera del ASE siguiente.
    /// </summary>
    internal static Destino ResolverTotalDeBloque(
        WorkbookPart parte,
        Worksheet hojaFisica,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        IReadOnlyList<DelimitadorBloquesHoja.BloqueHoja> bloques,
        int aseId,
        RolTotalBloque rol,
        string encabezado)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(hojaFisica);
        ArgumentNullException.ThrowIfNull(hoja);
        ArgumentNullException.ThrowIfNull(bloques);

        var bloque = bloques.FirstOrDefault(b => b.AseId == aseId)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{hoja.Nombre}' no declara el bloque del ASE {aseId} que exige el total {(rol == RolTotalBloque.Oportuno ? "oportuno" : "extemporáneo")}.");

        var fila = LocalizarFilaDelTotal(parte, hojaFisica, hoja, bloque, rol);
        var columna = ColumnaPorEncabezado(
            parte, hojaFisica, bloque,
            new Identidad(hoja.Nombre, aseId, NivelDestino.GranTotal, null, 1, encabezado));
        var referencia = $"{columna}{fila.ToString(CultureInfo.InvariantCulture)}";
        var celda = BuscarCelda(hojaFisica, referencia)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-SIN-SLOT: el total {(rol == RolTotalBloque.Oportuno ? "oportuno" : "extemporáneo")} del bloque ASE{aseId} resolvió a {referencia}, que no existe en '{hoja.Nombre}'.");

        return new Destino(fila, columna, referencia, celda);
    }

    private static int LocalizarFilaDelTotal(
        WorkbookPart parte,
        Worksheet hojaFisica,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        DelimitadorBloquesHoja.BloqueHoja bloque,
        RolTotalBloque rol)
    {
        var filas = RegionDeTotales(parte, hojaFisica, bloque);

        var oportunos = filas
            .Where(f => string.Equals(Texto(f, hoja.ColumnaGrupoResumen), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase)
                || (hoja.ColumnaGrupoResumen != "B"
                    && string.Equals(Texto(f, "B"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase)))
            .Where(f => string.Equals(Texto(f, hoja.ColumnaConceptoResumen), Oportuno, StringComparison.OrdinalIgnoreCase))
            .Where(f => string.Equals(Texto(f, "E"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase)
                || string.Equals(Texto(f, "C"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase)
                || string.Equals(Texto(f, "D"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase))
            .Select(f => (int)f.RowIndex!.Value)
            .ToList();

        if (oportunos.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{hoja.Nombre}' declara {oportunos.Count} filas de total oportuno en la región de totales del bloque ASE{bloque.AseId} ('{bloque.Rotulo}') y se exige exactamente una (rótulo '{DelimitadorBloquesHoja.MarcaTotal}' + '{Oportuno}' + '{DelimitadorBloquesHoja.MarcaTotal}'). El cuadro de totales se identifica por rótulo; nunca por posición.");
        }

        if (rol == RolTotalBloque.Oportuno)
        {
            return oportunos[0];
        }

        var filaOportuno = oportunos[0];
        var extemporaneos = filas
            .Where(f => f.RowIndex!.Value > filaOportuno)
            .TakeWhile(f => Texto(f, hoja.ColumnaGrupoResumen).Length == 0)
            .Where(f => string.Equals(Texto(f, hoja.ColumnaConceptoResumen), Extemporaneo, StringComparison.OrdinalIgnoreCase))
            .Where(f => string.Equals(Texto(f, "E"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase)
                || string.Equals(Texto(f, "C"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase)
                || string.Equals(Texto(f, "D"), DelimitadorBloquesHoja.MarcaTotal, StringComparison.OrdinalIgnoreCase))
            .Select(f => (int)f.RowIndex!.Value)
            .ToList();

        if (extemporaneos.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{hoja.Nombre}' declara {extemporaneos.Count} filas de total extemporáneo en el grupo del bloque ASE{bloque.AseId} ('{bloque.Rotulo}') y se exige exactamente una (primera '{Extemporaneo}' del grupo del total, antes del siguiente rótulo de grupo).");
        }

        return extemporaneos[0];

        string Texto(Row f, string columna) => LectorTextoHoja.TextoDeColumna(parte, f, columna).Trim();
    }

    /// <summary>
    /// Filas de la región de totales del bloque: desde la fila siguiente a su gran total hasta la
    /// cabecera del ASE siguiente (o el final de la hoja). Es estructura que la plantilla declara (el
    /// bloque y su cuadro), no un intervalo congelado.
    /// </summary>
    internal static IReadOnlyList<Row> RegionDeTotales(
        WorkbookPart parte,
        Worksheet hojaFisica,
        DelimitadorBloquesHoja.BloqueHoja bloque) =>
        hojaFisica.Elements<SheetData>().FirstOrDefault()?.Elements<Row>()
            .Where(r => r.RowIndex is not null && r.RowIndex!.Value > bloque.GranTotal)
            .OrderBy(r => r.RowIndex!.Value)
            .TakeWhile(r => !EsCabeceraDeBloque(parte, r))
            .ToList()
        ?? [];

    private const string Oportuno = "OPORTUNO";
    private const string Extemporaneo = "EXTEMPORANEO";

    private static bool EsCabeceraDeBloque(WorkbookPart parte, Row fila)
    {
        var texto = LectorTextoHoja.TextoDeColumna(parte, fila, "B").Trim();
        if (texto.Length == 0)
        {
            return false;
        }

        for (var id = 1; id <= Remuneracion.Core.Constants.CarpetasAse.Prefijos.Count; id++)
        {
            if (Remuneracion.Core.Models.SinonimosEmpresaR1.SonMismaEmpresa(
                    texto, Remuneracion.Core.Services.AseFactory.DesdeId(id).NombreCompleto))
            {
                return true;
            }
        }

        return false;
    }

    // ── Nivel tabla (CONSOLIDADO / AJUSTES - SF-T / DetRetri) ────────────────────────────────────
    /// <summary>
    /// Resuelve el destino dentro de una TABLA por ASE. Regla del T0 §3.1/H-59-3: el título vive en
    /// <c>F</c> y la tabla lo es porque DEBAJO declara el encabezado <c>B='ASE'</c> (la fila 141 del
    /// CONSOLIDADO repite el título 'R2 - ANTICIPOS A USUARIOS' como rótulo de columna y NO declara
    /// encabezado, por eso no es tabla). Sin título (DetRetri), el encabezado debe ser ÚNICO.
    /// </summary>
    internal static Destino ResolverTabla(
        WorkbookPart parte,
        Worksheet hojaFisica,
        IdentidadTabla identidad)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(hojaFisica);
        ArgumentNullException.ThrowIfNull(identidad);

        return ResolverTabla(parte, hojaFisica, identidad, porConcepto: false);
    }

    /// <summary>
    /// Igual que <see cref="ResolverTabla"/>, pero el rótulo de la columna se compara NORMALIZADO
    /// (sin acentos, sin espacios, sin decoraciones como <c>(*)</c>): la plantilla escribe el MISMO
    /// concepto con distinta decoración entre las tablas de las hojas de detalle —`'COSTO DE LA
    /// INTERVENTORÍA (*)'` en <c>DetRetri</c> y `'COSTO DE LA INTERVENTORÍA'` en <c>DetValiRetri</c>—
    /// y esa decoración no cambia el concepto al que la columna se refiere (§2-D2.3-bis: la
    /// identidad es el rótulo, nunca la celda). Sigue fallando con <c>ERR-SIN-SLOT</c> si el
    /// concepto no está declarado.
    /// </summary>
    internal static Destino ResolverTablaPorConcepto(
        WorkbookPart parte,
        Worksheet hojaFisica,
        IdentidadTabla identidad)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(hojaFisica);
        ArgumentNullException.ThrowIfNull(identidad);

        return ResolverTabla(parte, hojaFisica, identidad, porConcepto: true);
    }

    private static Destino ResolverTabla(
        WorkbookPart parte,
        Worksheet hojaFisica,
        IdentidadTabla identidad,
        bool porConcepto)
    {
        var filas = hojaFisica.Elements<SheetData>().FirstOrDefault()?
            .Elements<Row>()
            .Where(r => r.RowIndex is not null)
            .OrderBy(r => r.RowIndex!.Value)
            .ToList()
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{identidad.Hoja}' no tiene SheetData: no se puede resolver {identidad.Describir()}.");

        var encabezado = LocalizarEncabezadoTabla(parte, filas, identidad);
        var columna = ColumnaDeEncabezadoTabla(parte, hojaFisica, encabezado, identidad, porConcepto);
        var fila = LocalizarFilaTabla(parte, filas, encabezado, identidad);
        var referencia = $"{columna}{fila.ToString(CultureInfo.InvariantCulture)}";
        var celda = BuscarCelda(hojaFisica, referencia)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-SIN-SLOT: el destino {identidad.Describir()} se resolvió a la celda {referencia}, que no existe en la hoja '{identidad.Hoja}'. La plantilla no declara ese destino.");

        return new Destino(fila, columna, referencia, celda);
    }

    /// <summary>Filas del CONSOLIDADO que declaran un título de tabla en la columna <c>F</c>.</summary>
    internal static IReadOnlyList<int> FilasConTitulo(
        WorkbookPart parte,
        Worksheet hojaFisica,
        string titulo)
    {
        var filas = hojaFisica.Elements<SheetData>().FirstOrDefault()?.Elements<Row>() ?? [];
        return filas
            .Where(r => r.RowIndex is not null)
            .Where(r => string.Equals(
                LectorTextoHoja.TextoDeColumna(parte, r, "F").Trim(), titulo.Trim(), StringComparison.OrdinalIgnoreCase))
            .Select(r => (int)r.RowIndex!.Value)
            .ToList();
    }

    private static Row LocalizarEncabezadoTabla(
        WorkbookPart parte,
        IReadOnlyList<Row> filas,
        IdentidadTabla identidad)
    {
        var candidatos = filas
            .Where(r => string.Equals(
                LectorTextoHoja.TextoDeColumna(parte, r, "B").Trim(),
                EncabezadoTablaAse, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (identidad.Titulo is { } titulo)
        {
            var filaTitulo = filas.FirstOrDefault(r => string.Equals(
                LectorTextoHoja.TextoDeColumna(parte, r, "F").Trim(), titulo.Trim(), StringComparison.OrdinalIgnoreCase))
                ?? throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"ERR-PLANTILLA: la hoja '{identidad.Hoja}' no declara la tabla '{titulo}' (se esperaba su título en la columna F). Destino pedido: {identidad.Describir()}.");

            var encabezado = candidatos.FirstOrDefault(r => r.RowIndex!.Value > filaTitulo.RowIndex!.Value);
            if (encabezado is null)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"ERR-PLANTILLA: la tabla '{titulo}' de '{identidad.Hoja}' (título bajo la fila {filaTitulo.RowIndex}) no declara el encabezado '{EncabezadoTablaAse}' debajo: por la regla del T0 no es una tabla por ASE. Destino pedido: {identidad.Describir()}.");
            }

            return encabezado;
        }

        if (candidatos.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{identidad.Hoja}' declara {candidatos.Count} encabezados '{EncabezadoTablaAse}' (filas {(candidatos.Count == 0 ? "ninguna" : string.Join(", ", candidatos.Select(c => c.RowIndex!.Value)))}) y el destino {identidad.Describir()} exige exactamente uno para identificarse sin ambigüedad.");
        }

        return candidatos[0];
    }

    private static string ColumnaDeEncabezadoTabla(
        WorkbookPart parte,
        Worksheet hojaFisica,
        Row encabezado,
        IdentidadTabla identidad,
        bool porConcepto = false)
    {
        foreach (var celda in encabezado.Elements<Cell>())
        {
            var letra = LetraDeColumna(celda.CellReference?.Value);
            if (letra.Length == 0)
            {
                continue;
            }

            var texto = LectorTextoHoja.Texto(parte, celda).Trim();
            var coincide = porConcepto
                ? string.Equals(
                    DetallePeriodoDerivado.NormalizarRotulo(texto),
                    DetallePeriodoDerivado.NormalizarRotulo(identidad.Encabezado),
                    StringComparison.Ordinal)
                : string.Equals(texto, identidad.Encabezado.Trim(), StringComparison.OrdinalIgnoreCase);
            if (coincide)
            {
                return letra;
            }
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-SIN-SLOT: el encabezado de la tabla de '{identidad.Hoja}' no declara una columna '{identidad.Encabezado}'. Destino pedido: {identidad.Describir()}. Encabezados declarados: {string.Join(", ", encabezado.Elements<Cell>().Select(c => LectorTextoHoja.Texto(parte, c).Trim()).Where(t => t.Length > 0))}.");
    }

    private static int LocalizarFilaTabla(
        WorkbookPart parte,
        IReadOnlyList<Row> filas,
        Row encabezado,
        IdentidadTabla identidad)
    {
        foreach (var fila in filas.Where(r => r.RowIndex!.Value > encabezado.RowIndex!.Value))
        {
            if (identidad.AseId is { } aseId)
            {
                var textoB = LectorTextoHoja.TextoDeColumna(parte, fila, "B").Trim();
                if (int.TryParse(textoB, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) && id == aseId)
                {
                    return (int)fila.RowIndex!.Value;
                }

                continue;
            }

            // Fila de cierre de la tabla: la marca TOTAL vive en la columna C y su B va vacía.
            var textoC = LectorTextoHoja.TextoDeColumna(parte, fila, "C").Trim();
            if (string.Equals(textoC, MarcaTotalTabla, StringComparison.OrdinalIgnoreCase))
            {
                return (int)fila.RowIndex!.Value;
            }
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-SIN-SLOT: la tabla de '{identidad.Hoja}' no declara {(identidad.AseId is { } ase ? $"la fila del ASE {ase} (columna B = id numérico)" : $"la fila de cierre (columna C = '{MarcaTotalTabla}')")} bajo el encabezado de la fila {encabezado.RowIndex}. Destino pedido: {identidad.Describir()}.");
    }

    // ── Utilidades de referencia (sin literales de fila/celda) ───────────────────────────────────

    private static CalculoInvalidoException SinDestino(
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        DelimitadorBloquesHoja.BloqueHoja bloque,
        Identidad identidad) =>
        new(
            CodigoError.Plantilla,
            $"ERR-SIN-SLOT: el destino {identidad.Describir()} no existe en la hoja '{hoja.Nombre}' dentro del bloque '{bloque.Rotulo}' ({bloque.Cabecera}-{bloque.GranTotal}). La plantilla no declara ese cierre: no se reacomoda a otra fila ni se coloca por posición.");

    /// <summary>Celda del worksheet por referencia ya calculada (nunca una referencia literal).</summary>
    internal static Cell? BuscarCelda(Worksheet hoja, string referencia) =>
        hoja.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));

    /// <summary>Letra(s) de columna de una referencia A1 (sin dígitos ni <c>$</c>).</summary>
    internal static string LetraDeColumna(string? referencia)
    {
        if (string.IsNullOrEmpty(referencia))
        {
            return string.Empty;
        }

        var fin = 0;
        while (fin < referencia.Length && (char.IsLetter(referencia[fin]) || referencia[fin] == '$'))
        {
            fin++;
        }

        return referencia[..fin].Replace("$", string.Empty, StringComparison.Ordinal);
    }
}
