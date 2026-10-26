using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 59 (H1.1, D-1, S2.1/S2.2): DELIMITADOR de bloques de ASE leídos de la PROPIA hoja.
///
/// Problema que elimina (Plan 59 §0.4-f/§0.4-g): la app localizaba el bloque con ventanas de filas
/// congeladas por ASE (el mapa `1:(6,72) … 5:(603,676)` que este H retiró); con la plantilla v3 el
/// cierre <c>Subs/Cont/Total</c> de ASE1 quedó en la fila 85, fuera de la ventana 6‑72, y el escritor
/// reportó <c>ERR-SIN-SLOT</c> sobre un destino que la plantilla SÍ declara.
///
/// Regla de delimitación (cero filas conocidas, cero ramas por ASE/empresa/período/quincena):
/// <list type="bullet">
/// <item><b>Inicio</b>: la fila cuyo rótulo (columna <c>B</c>) es un ASE del catálogo de dominio
/// (<see cref="AseFactory"/>) comparado con el catálogo abierto de sinonimia
/// (<see cref="SinonimosEmpresaR1.SonMismaEmpresa"/>). Es la CABECERA DE BLOQUE ratificada (D‑2):
/// no se usa «fila siguiente al gran total anterior» porque ese hueco contiene el cuadro‑resumen,
/// que reutiliza las firmas de cierre (T0 §2.5: en R2 <c>C='OPORTUNO'/D='Total'</c> imita un cierre
/// de canal y en R4 <c>B='OPORTUNO'/C='Total'</c> imita un cierre de sección);</item>
/// <item><b>Fin</b>: la PRIMERA fila posterior a la cabecera que declara el gran total del bloque
/// (<c>A='Total'</c> con <c>B</c> vacía: cierre de bloque). El cuadro‑resumen que va debajo queda
/// FUERA del rango por construcción;</item>
/// <item><b>Clasificación por FIRMA, no por posición</b>: cada fila del rango se clasifica por su
/// combinación de etiquetas contra la gramática de cierres de SU hoja (ver
/// <see cref="HojaDelimitable"/>). No hay «cabecera + n» ni ningún offset aritmético: la fila que
/// no declara firma de cierre y no es la fila de encabezados de columna es, por definición, fila de
/// datos.</item>
/// </list>
///
/// Vocabulario de cierres por hoja (T0 §2, medido por bytes; TODOS son TRAILER — cierran lo que
/// está arriba, nunca encabezan). Se declara como DATOS de la gramática, nunca como código con
/// ramas por hoja dentro de la delimitación:
/// <list type="bullet">
/// <item><b>R1</b> (4 niveles): canal <c>D=&lt;canal&gt;, E='Total'</c>; empresa <c>C=&lt;empresa&gt;,
/// D='Total'</c>; subsección <c>B=&lt;etiqueta&gt;, C='Total'</c>; sección <c>A=&lt;etiqueta&gt;,
/// B='Total'</c>; gran total <c>A='Total'</c>;</item>
/// <item><b>R2</b> (3 niveles, sin subsección): canal <c>C/D='Total'</c>; empresa <c>B/C='Total'</c>;
/// sección <c>A/B='Total'</c>; gran total <c>A='Total'</c>;</item>
/// <item><b>R4</b> (2 niveles, sin canal ni subsección): empresa <c>A/B='Total'</c>; sección
/// <c>B/C='Total'</c>; gran total <c>A='Total'</c>.</item>
/// </list>
///
/// Los rótulos de sección NO son allow‑list de negocio: el cierre de sección se reconoce por su
/// FORMA (<c>A=&lt;algo&gt;/B='Total'</c> con el resto vacío), no por su texto. La única lista
/// cerrada es la de la GRAMÁTICA (<c>Total</c> como marca de cierre) y la del catálogo de dominio
/// de ASE.
///
/// Fail‑fast accionable (G‑TRAZA): bloque sin gran total, cierres incompletos, fila que encaja en dos
/// niveles o bloque de un ASE del catálogo ausente/duplicado ⇒ <c>ERR-PLANTILLA</c> con hoja, rótulo
/// esperado y filas halladas. Nunca remapeo por posición.
/// </summary>
internal static class DelimitadorBloquesHoja
{
    /// <summary>Nombre de la hoja de recaudo por componentes (R1).</summary>
    internal const string NombreHojaR1 = "Reporte Componentes R1";

    /// <summary>Nombre de la hoja de anticipos a usuarios (R2).</summary>
    internal const string NombreHojaR2 = "Rem. Anticipos R2";

    /// <summary>Nombre de la hoja de reversión de pagos (R4).</summary>
    internal const string NombreHojaR4 = "Reversion Pagos R4";

    // ── Gramática de COLUMNAS (roles de la firma A–E, no coordenadas de fila) ────────────────────
    private const string ColumnaSeccion = "A";
    private const string ColumnaSubseccion = "B";
    private const string ColumnaFuente = "C";
    private const string ColumnaCanal = "D";
    private const string ColumnaConcepto = "E";

    /// <summary>Marca literal de cierre (única palabra cerrada de la gramática).</summary>
    private const string MarcaTotal = "Total";

    /// <summary>
    /// Textos mínimos en las columnas de componente para reconocer la fila de encabezados de columna
    /// (esa fila declara varios rótulos y ningún valor numérico; una fila de datos declara valores).
    /// </summary>
    private const int TextosMinimosEncabezado = 2;

    /// <summary>
    /// Columnas del interior con encabezado de componente (F..Z), en orden físico. Gramática del
    /// contrato de columnas por encabezado (Plan 44 F1), compartida con el escritor values‑only.
    /// </summary>
    internal static readonly string[] ColumnasComponentes =
    [
        "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
    ];

    /// <summary>Nivel de cierre que declara una fila (vocabulario del T0, por hoja).</summary>
    internal enum NivelCierreHoja
    {
        /// <summary>Gran total del bloque (<c>A='Total'</c> con B vacía): CIERRA el bloque.</summary>
        GranTotal,

        /// <summary>Cierre de sección (<c>A=&lt;sección&gt;/B='Total'</c>).</summary>
        Seccion,

        /// <summary>Cierre de subsección/cuadro (<c>B=&lt;subsección&gt;/C='Total'</c>). Solo R1.</summary>
        Subseccion,

        /// <summary>Cierre del grupo de canales de una empresa/fuente (trailer).</summary>
        EmpresaBloque,

        /// <summary>Cierre de canal (<c>D=&lt;canal&gt;/E='Total'</c> en R1, <c>C/D='Total'</c> en R2).</summary>
        Canal,
    }

    /// <summary>Papel de una fila dentro del bloque ya delimitado.</summary>
    internal enum TipoFilaBloqueHoja
    {
        /// <summary>Cabecera del bloque: la fila cuyo rótulo (columna B) es el ASE.</summary>
        CabeceraBloque,

        /// <summary>Fila de encabezados de columna de componente (declara rótulos, no valores).</summary>
        EncabezadoColumnas,

        /// <summary>Fila de datos (conceptos): la que no declara cierre ni encabezado.</summary>
        Dato,

        /// <summary>Cierre de canal.</summary>
        CierreCanal,

        /// <summary>Cierre del grupo de una empresa/fuente.</summary>
        CierreEmpresa,

        /// <summary>Cierre de subsección/cuadro.</summary>
        CierreSubseccion,

        /// <summary>Cierre de sección.</summary>
        CierreSeccion,

        /// <summary>Gran total del bloque.</summary>
        GranTotal,
    }

    /// <summary>
    /// Firma de cierre de una hoja: <paramref name="ColumnaEtiqueta"/> trae la etiqueta
    /// (vacía ⇒ no es cierre), <paramref name="ColumnaMarca"/> trae la marca <c>'Total'</c> (cadena
    /// vacía ⇒ el nivel no exige marca, caso del gran total) y <paramref name="ColumnasVacias"/> son
    /// las columnas que deben estar vacías para que la forma sea inequívoca.
    /// <paramref name="ValorEtiqueta"/> fija el valor exigido en la columna de etiqueta (solo el gran
    /// total lo usa: <c>'Total'</c>); cuando es <c>null</c> la etiqueta debe ser no vacía y distinta
    /// de la marca.
    /// </summary>
    internal sealed record FirmaCierreHoja(
        NivelCierreHoja Nivel,
        string ColumnaEtiqueta,
        string ColumnaMarca,
        IReadOnlyList<string> ColumnasVacias,
        string? ValorEtiqueta = null);

    /// <summary>
    /// Hoja delimitable: su nombre + el vocabulario de cierres que ELLA declara. La cardinalidad de
    /// niveles (4/3/2) es la del T0, no una rama de código de la delimitación.
    /// </summary>
    internal sealed record HojaDelimitable(string Nombre, IReadOnlyList<FirmaCierreHoja> Cierres)
    {
        /// <summary>Verdad si la hoja declara dimensión de subsección (R1).</summary>
        internal bool DeclaraSubseccion => Cierres.Any(c => c.Nivel == NivelCierreHoja.Subseccion);

        /// <summary>Verdad si la hoja declara dimensión de canal (R1/R2; R4 no la tiene).</summary>
        internal bool DeclaraCanal => Cierres.Any(c => c.Nivel == NivelCierreHoja.Canal);
    }

    /// <summary>Fila clasificada del bloque, con las etiquetas A–E leídas del propio workbook.</summary>
    internal sealed record FilaBloqueHoja(int Fila, TipoFilaBloqueHoja Tipo, FilaEspejoR1 Etiquetas)
    {
        /// <summary>
        /// Verdad si la fila participa del contrato de slots (todo menos la cabecera de bloque y la
        /// fila de encabezados de columna).
        /// </summary>
        internal bool EsSlot => Tipo is not (TipoFilaBloqueHoja.CabeceraBloque or TipoFilaBloqueHoja.EncabezadoColumnas);
    }

    /// <summary>
    /// Bloque delimitado: identidad (ASE del catálogo + rótulo leído), rango
    /// <see cref="Cabecera"/>..<see cref="GranTotal"/> (inclusivo) y las filas clasificadas.
    /// </summary>
    internal sealed record BloqueHoja(
        int AseId,
        string Rotulo,
        int Cabecera,
        int GranTotal,
        int EncabezadoColumnas,
        IReadOnlyList<FilaBloqueHoja> Filas)
    {
        /// <summary>Filas que el contrato de slots debe considerar (datos + cierres).</summary>
        internal IReadOnlyList<FilaBloqueHoja> FilasSlot => Filas.Where(f => f.EsSlot).ToList();

        /// <summary>Cuántas filas del bloque declaran el nivel de cierre indicado.</summary>
        internal int Cuenta(TipoFilaBloqueHoja tipo) => Filas.Count(f => f.Tipo == tipo);

        /// <summary>Filas del bloque que declaran el nivel de cierre indicado, en orden físico.</summary>
        internal IReadOnlyList<FilaBloqueHoja> Cierres(TipoFilaBloqueHoja tipo) =>
            Filas.Where(f => f.Tipo == tipo).ToList();
    }

    // ── Vocabulario de cierres por hoja (T0 §2, vinculante) ──────────────────────────────────────

    /// <summary>R1: 4 niveles (canal, empresa, subsección, sección) + gran total.</summary>
    internal static readonly HojaDelimitable R1 = new(
        NombreHojaR1,
        [
            new(
                NivelCierreHoja.GranTotal, ColumnaSeccion, string.Empty,
                [ColumnaSubseccion, ColumnaFuente, ColumnaCanal, ColumnaConcepto], MarcaTotal),
            new(
                NivelCierreHoja.Seccion, ColumnaSeccion, ColumnaSubseccion,
                [ColumnaFuente, ColumnaCanal, ColumnaConcepto]),
            new(
                NivelCierreHoja.Subseccion, ColumnaSubseccion, ColumnaFuente,
                [ColumnaSeccion]),
            new(
                NivelCierreHoja.EmpresaBloque, ColumnaFuente, ColumnaCanal,
                [ColumnaSeccion, ColumnaSubseccion]),
            new(
                NivelCierreHoja.Canal, ColumnaCanal, ColumnaConcepto,
                [ColumnaSeccion, ColumnaSubseccion, ColumnaFuente]),
        ]);

    /// <summary>R2: 3 niveles (canal, empresa, sección) + gran total; sin subsección.</summary>
    internal static readonly HojaDelimitable R2 = new(
        NombreHojaR2,
        [
            new(
                NivelCierreHoja.GranTotal, ColumnaSeccion, string.Empty,
                [ColumnaSubseccion], MarcaTotal),
            new(
                NivelCierreHoja.Seccion, ColumnaSeccion, ColumnaSubseccion,
                [ColumnaFuente, ColumnaCanal]),
            new(
                NivelCierreHoja.EmpresaBloque, ColumnaSubseccion, ColumnaFuente,
                [ColumnaSeccion]),
            new(
                NivelCierreHoja.Canal, ColumnaFuente, ColumnaCanal,
                [ColumnaSeccion, ColumnaSubseccion]),
        ]);

    /// <summary>R4: 2 niveles (empresa, sección) + gran total; sin canal ni subsección.</summary>
    internal static readonly HojaDelimitable R4 = new(
        NombreHojaR4,
        [
            new(
                NivelCierreHoja.GranTotal, ColumnaSeccion, string.Empty,
                [ColumnaSubseccion], MarcaTotal),
            new(
                NivelCierreHoja.EmpresaBloque, ColumnaSeccion, ColumnaSubseccion,
                []),
            new(
                NivelCierreHoja.Seccion, ColumnaSubseccion, ColumnaFuente,
                [ColumnaSeccion]),
        ]);

    /// <summary>
    /// Delimita TODOS los bloques de la hoja (uno por ASE del catálogo de dominio, en orden físico)
    /// leyendo solo lo que la hoja declara. Falla accionablemente si la estructura de cierres está
    /// incompleta o si un ASE del catálogo no declara bloque.
    /// </summary>
    internal static IReadOnlyList<BloqueHoja> Delimitar(
        WorkbookPart parte,
        HojaDelimitable hoja,
        Worksheet? worksheet = null)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(hoja);

        var hojaFisica = worksheet ?? ResolverHoja(parte, hoja.Nombre);
        var tabla = hojaFisica.Elements<SheetData>().FirstOrDefault()
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La hoja '{hoja.Nombre}' no tiene SheetData: no se pueden delimitar sus bloques por la cabecera de bloque (columna {ColumnaSubseccion} = rótulo del ASE) ni por el gran total declarado.");

        var filas = tabla.Elements<Row>()
            .Where(f => f.RowIndex is not null)
            .OrderBy(f => f.RowIndex!.Value)
            .ToList();
        var etiquetas = filas.Select(f => LeerEtiquetas(parte, f)).ToList();

        // Cabeceras de bloque: la fila cuyo rótulo (columna B) es un ASE del catálogo de dominio.
        var cabeceras = new List<int>();
        for (var i = 0; i < filas.Count; i++)
        {
            if (TryAseId(etiquetas[i].B, out _))
            {
                cabeceras.Add(i);
            }
        }

        if (cabeceras.Count == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La hoja '{hoja.Nombre}' no declara ninguna cabecera de bloque: ninguna fila trae en la columna {ColumnaSubseccion} un rótulo del catálogo de ASE ({RotulosDelCatalogo()}). La delimitación es por etiqueta, nunca por posición.");
        }

        var bloques = new List<BloqueHoja>(cabeceras.Count);
        for (var k = 0; k < cabeceras.Count; k++)
        {
            var desde = cabeceras[k];
            var hastaExclusivo = k + 1 < cabeceras.Count ? cabeceras[k + 1] : filas.Count;
            TryAseId(etiquetas[desde].B, out var aseId);
            bloques.Add(ConstruirBloque(parte, hoja, filas, etiquetas, desde, hastaExclusivo, aseId));
        }

        ValidarCoberturaDelCatalogo(hoja, bloques);
        return bloques;
    }

    /// <summary>
    /// Delimita la hoja y devuelve el bloque del ASE pedido, verificando IDENTIDAD POR RÓTULO contra
    /// el catálogo de dominio (nunca por índice de fila). Es la verificación que sustituye a
    /// <c>VerificarCabeceraBloque</c> del camino con ventana congelada.
    /// </summary>
    internal static BloqueHoja DelimitarBloque(
        WorkbookPart parte,
        HojaDelimitable hoja,
        int aseId,
        Worksheet? worksheet = null)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(hoja);

        var esperado = AseFactory.DesdeId(aseId).NombreCompleto;
        var bloques = Delimitar(parte, hoja, worksheet);
        var bloque = bloques.FirstOrDefault(b => b.AseId == aseId);
        if (bloque is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La hoja '{hoja.Nombre}' no declara el bloque del ASE {aseId}: se esperaba una cabecera con el rótulo '{esperado}' en la columna {ColumnaSubseccion} y un gran total ({ColumnaSeccion}='{MarcaTotal}' con {ColumnaSubseccion} vacía) debajo. Bloques hallados: {Describir(bloques)}. Fail-fast antes de escribir: la identidad del bloque es por rótulo del catálogo, nunca por fila.");
        }

        if (!SinonimosEmpresaR1.SonMismaEmpresa(bloque.Rotulo, esperado))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El bloque del ASE {aseId} de la hoja '{hoja.Nombre}' declara el rótulo '{bloque.Rotulo}' (fila {bloque.Cabecera}) y no corresponde a '{esperado}' del catálogo de dominio. Fail-fast antes de escribir.");
        }

        return bloque;
    }

    // ── Construcción del bloque ──────────────────────────────────────────────────────────────────

    private static BloqueHoja ConstruirBloque(
        WorkbookPart parte,
        HojaDelimitable hoja,
        IReadOnlyList<Row> filas,
        IReadOnlyList<FilaEspejoR1> etiquetas,
        int desde,
        int hastaExclusivo,
        int aseId)
    {
        var rotulo = etiquetas[desde].B.Trim();
        var cabecera = Numero(filas[desde]);

        // FIN del bloque: la PRIMERA fila posterior a la cabecera que declara el gran total. El
        // cuadro-resumen que vive debajo reutiliza firmas de cierre en R2/R4 (T0 §2.5), por eso el
        // rango se corta aquí y no en la cabecera del ASE siguiente.
        var indiceTotal = -1;
        for (var i = desde + 1; i < hastaExclusivo; i++)
        {
            if (NivelDe(hoja, etiquetas[i], Numero(filas[i])) == NivelCierreHoja.GranTotal)
            {
                indiceTotal = i;
                break;
            }
        }

        if (indiceTotal < 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El bloque del ASE {aseId} ('{rotulo}', cabecera en la fila {cabecera}) de la hoja '{hoja.Nombre}' NO declara su gran total ({ColumnaSeccion}='{MarcaTotal}' con {ColumnaSubseccion} vacía) antes de la cabecera siguiente ni del final de la hoja. Estructura de cierres incompleta: no se delimita por posición.");
        }

        var clasificadas = new List<FilaBloqueHoja>(indiceTotal - desde + 1);
        var encabezados = new List<int>();
        for (var i = desde; i <= indiceTotal; i++)
        {
            var numero = Numero(filas[i]);
            TipoFilaBloqueHoja tipo;
            if (i == desde)
            {
                tipo = TipoFilaBloqueHoja.CabeceraBloque;
            }
            else
            {
                tipo = NivelDe(hoja, etiquetas[i], numero) switch
                {
                    NivelCierreHoja.GranTotal => TipoFilaBloqueHoja.GranTotal,
                    NivelCierreHoja.Seccion => TipoFilaBloqueHoja.CierreSeccion,
                    NivelCierreHoja.Subseccion => TipoFilaBloqueHoja.CierreSubseccion,
                    NivelCierreHoja.EmpresaBloque => TipoFilaBloqueHoja.CierreEmpresa,
                    NivelCierreHoja.Canal => TipoFilaBloqueHoja.CierreCanal,
                    null => EsEncabezadoColumnas(filas[i])
                        ? TipoFilaBloqueHoja.EncabezadoColumnas
                        : TipoFilaBloqueHoja.Dato,
                    _ => TipoFilaBloqueHoja.Dato,
                };
                if (tipo == TipoFilaBloqueHoja.EncabezadoColumnas)
                {
                    encabezados.Add(numero);
                }
            }

            clasificadas.Add(new FilaBloqueHoja(numero, tipo, etiquetas[i]));
        }

        if (encabezados.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El bloque del ASE {aseId} ('{rotulo}') de la hoja '{hoja.Nombre}' declara {encabezados.Count} filas de encabezados de columna ({(encabezados.Count == 0 ? "ninguna" : string.Join(", ", encabezados))}) en {cabecera}-{Numero(filas[indiceTotal])}: se exige exactamente una fila que declare los rótulos de las columnas de componente ({string.Join(",", ColumnasComponentes)}) sin valores numéricos. La primera fila de datos no se calcula por offset.");
        }

        var granTotal = Numero(filas[indiceTotal]);
        if (clasificadas.All(f => f.Tipo != TipoFilaBloqueHoja.CierreSeccion))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El bloque del ASE {aseId} ('{rotulo}') de la hoja '{hoja.Nombre}' ({cabecera}-{granTotal}) no declara ningún cierre de sección ({hoja.Cierres.First(c => c.Nivel == NivelCierreHoja.Seccion).ColumnaEtiqueta}=<sección> con {hoja.Cierres.First(c => c.Nivel == NivelCierreHoja.Seccion).ColumnaMarca}='{MarcaTotal}'): la estructura de cierres está incompleta y el bloque no se puede delimitar.");
        }

        if (hoja.DeclaraSubseccion && clasificadas.All(f => f.Tipo != TipoFilaBloqueHoja.CierreSubseccion))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El bloque del ASE {aseId} ('{rotulo}') de la hoja '{hoja.Nombre}' ({cabecera}-{granTotal}) no declara ningún cierre de subsección ({ColumnaSubseccion}=<subsección> con {ColumnaFuente}='{MarcaTotal}'): la estructura de cierres está incompleta.");
        }

        return new BloqueHoja(aseId, rotulo, cabecera, granTotal, encabezados[0], clasificadas);
    }

    /// <summary>
    /// Nivel de cierre que declara la fila según la gramática de SU hoja, o <c>null</c> si no cierra
    /// nada (fila de dato o de encabezado). Una fila que encaja en DOS niveles es una plantilla
    /// ambigua ⇒ fail-fast (nunca se elige uno por orden).
    /// </summary>
    private static NivelCierreHoja? NivelDe(HojaDelimitable hoja, FilaEspejoR1 fila, int numeroFila)
    {
        NivelCierreHoja? encontrado = null;
        foreach (var firma in hoja.Cierres)
        {
            if (!Encaja(firma, fila))
            {
                continue;
            }

            if (encontrado is { } previo)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La hoja '{hoja.Nombre}' fila {numeroFila} declara una firma ambigua A|B|C|D|E='{fila.Firma}': encaja a la vez en el cierre '{previo}' y en el cierre '{firma.Nivel}'. La gramática de cierres de la hoja no permite desambiguar por etiqueta.");
            }

            encontrado = firma.Nivel;
        }

        return encontrado;
    }

    private static bool Encaja(FirmaCierreHoja firma, FilaEspejoR1 fila)
    {
        var etiqueta = Valor(fila, firma.ColumnaEtiqueta);
        if (firma.ValorEtiqueta is { } exigido)
        {
            if (!string.Equals(etiqueta, exigido, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }
        else if (etiqueta.Length == 0 || string.Equals(etiqueta, MarcaTotal, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (firma.ColumnaMarca.Length > 0
            && !string.Equals(Valor(fila, firma.ColumnaMarca), MarcaTotal, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        foreach (var columna in firma.ColumnasVacias)
        {
            if (Valor(fila, columna).Length > 0)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>Etiqueta de la columna indicada dentro de la firma A–E (rol de columna, no coordenada).</summary>
    private static string Valor(FilaEspejoR1 fila, string columna) => columna switch
    {
        ColumnaSeccion => fila.A.Trim(),
        ColumnaSubseccion => fila.B.Trim(),
        ColumnaFuente => fila.C.Trim(),
        ColumnaCanal => fila.D.Trim(),
        ColumnaConcepto => fila.E.Trim(),
        _ => string.Empty,
    };

    /// <summary>
    /// Verdad si la fila es el encabezado de columnas del bloque: declara rótulos (texto) en las
    /// columnas de componente y NINGÚN valor numérico. Estructural: no depende de la posición ni del
    /// contenido concreto de los rótulos.
    /// </summary>
    private static bool EsEncabezadoColumnas(Row fila)
    {
        var textos = 0;
        foreach (var celda in fila.Elements<Cell>())
        {
            if (!EsColumnaDeComponente(ColumnaDe(celda.CellReference?.Value)))
            {
                continue;
            }

            var dataType = celda.DataType?.Value;
            if (dataType == CellValues.SharedString || dataType == CellValues.InlineString)
            {
                textos++;
                continue;
            }

            if (!string.IsNullOrWhiteSpace(celda.CellValue?.InnerText))
            {
                // Un valor numérico en la zona de componentes ⇒ es fila de datos, no de encabezados.
                return false;
            }
        }

        return textos >= TextosMinimosEncabezado;
    }

    private static bool EsColumnaDeComponente(string columna) =>
        columna.Length > 0
        && ColumnasComponentes.Contains(columna, StringComparer.OrdinalIgnoreCase);

    /// <summary>Letra(s) de columna de una referencia A1 (sin dígitos ni <c>$</c>).</summary>
    private static string ColumnaDe(string? referencia)
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

    private static FilaEspejoR1 LeerEtiquetas(WorkbookPart parte, Row fila) => new()
    {
        A = LectorTextoHoja.TextoDeColumna(parte, fila, ColumnaSeccion),
        B = LectorTextoHoja.TextoDeColumna(parte, fila, ColumnaSubseccion),
        C = LectorTextoHoja.TextoDeColumna(parte, fila, ColumnaFuente),
        D = LectorTextoHoja.TextoDeColumna(parte, fila, ColumnaCanal),
        E = LectorTextoHoja.TextoDeColumna(parte, fila, ColumnaConcepto),
    };

    private static bool TryAseId(string? rotulo, out int aseId)
    {
        aseId = 0;
        if (string.IsNullOrWhiteSpace(rotulo))
        {
            return false;
        }

        for (var id = 1; id <= CarpetasAse.Prefijos.Count; id++)
        {
            if (SinonimosEmpresaR1.SonMismaEmpresa(rotulo, AseFactory.DesdeId(id).NombreCompleto))
            {
                aseId = id;
                return true;
            }
        }

        return false;
    }

    private static void ValidarCoberturaDelCatalogo(HojaDelimitable hoja, IReadOnlyList<BloqueHoja> bloques)
    {
        for (var id = 1; id <= CarpetasAse.Prefijos.Count; id++)
        {
            var delAse = bloques.Where(b => b.AseId == id).ToList();
            if (delAse.Count == 0)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La hoja '{hoja.Nombre}' no declara bloque para el ASE {id}: se esperaba una cabecera con el rótulo '{AseFactory.DesdeId(id).NombreCompleto}' en la columna {ColumnaSubseccion}. Bloques hallados: {Describir(bloques)}.");
            }

            if (delAse.Count > 1)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La hoja '{hoja.Nombre}' declara {delAse.Count} bloques para el ASE {id} (filas {string.Join(", ", delAse.Select(b => b.Cabecera))}): la identidad del bloque debe ser única. La delimitación es por rótulo, nunca por posición.");
            }
        }
    }

    private static string Describir(IReadOnlyList<BloqueHoja> bloques) =>
        bloques.Count == 0
            ? "(ninguno)"
            : string.Join(", ", bloques.Select(b => $"ASE {b.AseId} '{b.Rotulo}' ({b.Cabecera}-{b.GranTotal})"));

    private static string RotulosDelCatalogo() =>
        string.Join(", ", Enumerable.Range(1, CarpetasAse.Prefijos.Count)
            .Select(id => $"ASE {id}='{AseFactory.DesdeId(id).NombreCompleto}'"));

    private static int Numero(Row fila) => (int)fila.RowIndex!.Value;

    private static Worksheet ResolverHoja(WorkbookPart parte, string nombreHoja)
    {
        var hoja = parte.Workbook?.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El workbook no trae la hoja '{nombreHoja}' para delimitar sus bloques de ASE.");
        return ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La hoja '{nombreHoja}' no tiene WorksheetPart válido para delimitar sus bloques de ASE.");
    }
}
