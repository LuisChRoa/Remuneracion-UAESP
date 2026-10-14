using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 36 (T1, R-B-1/R-B-2/R-B-4 — D-A/D-B/D-F): escritor values-only del interior de
/// <c>Reporte Componentes R1</c> sobre la plantilla estandarizada única.
///
/// Doctrina slot: la app NUNCA inserta/borra filas y NUNCA escribe <c>&lt;f&gt;</c>; solo pega
/// VALORES (<c>&lt;v&gt;</c>) en el slot de datos que corresponde por firma + frontera + sinonimia;
/// lo no entregado queda en el cero del template; la fuente que trae empresa/fila sin slot hace
/// fail-fast <c>ERR-SIN-SLOT</c> (nunca remap silencioso, nunca insert).
///
/// Resolución (arsenal reinvertido, cero lógica nueva de negocio):
/// <list type="bullet">
/// <item>firma como localizador (<see cref="R1FirmaInterior"/> — mismos predicados que hoy);</item>
/// <item>tipo por frontera (<see cref="R1FirmaInterior.EsAplicacionPorFrontera"/>, Plan 35-T2 intacto),
/// exigido solo en filas-dato de empresa (<see cref="R1FirmaInterior.EsDatoEmpresa(FilaEspejoR1)"/>)
/// y aplicado como preferencia con fallback a firma sola cuando la clasificación del lado fuente
/// degenera (fuente sin fila <c>Aplicacion</c>);</item>
/// <item>empresa por sinonimia (<see cref="SinonimosEmpresaR1.SonMismaEmpresa"/>, Plan 34 intacto,
/// catálogo abierto — cero ramas por nombre de empresa);</item>
/// <item>AFaseo dual-firma (R-T-2): la fuente en variante <c>A='AFaseo' B='Total'</c> o en variante
/// <c>D='AFASEO' E='Total'</c> resuelve al mismo slot F0 del ASE.</item>
/// </list>
///
/// Contrato de valores (W1 del review de T1): el slot recibe el DECIMAL EXACTO de la fuente (sin
/// redondeo); el redondeo a entero es exclusivo del <c>DetRetri</c> (Plan 29 Unidad D) y no se
/// aplica aquí. Ver la justificación completa en <see cref="EscribirValor"/>.
///
/// Emparejamiento por FIRMA con cursor INDEPENDIENTE por firma (refinamiento T1 del Plan 36,
/// hallazgo A de T4): en la plantilla estandarizada los slots-resumen (AFaseo, Mes, …) están
/// INTERCALADOS dentro del bloque y la fuente puede emitirlos en un orden distinto — p. ej. el
/// AFaseo del ASE2 vive en F121 (en medio del bloque) mientras la fuente lo emite DESPUÉS de
/// consumir filas posteriores del template. Un cursor GLOBAL monotónico no puede retroceder y deja
/// ese AFaseo sin slot (falso <c>ERR-SIN-SLOT</c>). Por eso cada fila fuente consume el próximo
/// slot libre DENTRO DE SU PROPIA FIRMA (lista ordenada de slots con firma compatible): un slot
/// intercalado antes en el template puede ser consumido por una fila emitida después, y cada firma
/// avanza su propio cursor sin interferir con las demás. Los slots no consumidos quedan intactos
/// en cero. Guard pre-write (REGLA, no excepción): la celda destino debe existir y NO traer
/// <c>&lt;f&gt;</c>; si trae fórmula → <c>ERR-PLANTILLA</c> con hoja + celda + slot. Cero
/// <c>InsertRow</c>/<c>DeleteRow</c>, cero <c>new CellFormula</c>, cero Reanclar en este path (el
/// grep de cierre T1 lo demuestra).
///
/// El writer NO recibe ningún período (D-H) y no nombra ninguna empresa (catálogo abierto).
/// <see cref="CompositorInteriorR1"/> queda como referencia de firma, no se usa aquí (D-C).
/// </summary>
internal static class EscritorValoresSlotR1
{
    public const string HojaR1 = "Reporte Componentes R1";

    /// <summary>
    /// Plan 36 (T0): rango de filas de DATOS (slots) del bloque R1 por ASE en la plantilla
    /// estandarizada única, 1-based e inclusivo. Congelado desde los bytes del disco por
    /// <c>ContratoSlotR1T0Tests</c> (cualquier deriva del template rompe T0 antes que esto):
    /// ASE1 6-72, ASE2 110-210, ASE3 289-355, ASE4 450-550, ASE5 603-676. Geometría del contrato,
    /// no literales de período (D-H intacta: ningún período entra ni sale de este path).
    /// </summary>
    internal static (int Inicio, int Fin) RangoDatosAse(int aseId) => aseId switch
    {
        1 => (6, 72),
        2 => (110, 210),
        3 => (289, 355),
        4 => (450, 550),
        5 => (603, 676),
        _ => throw new ArgumentOutOfRangeException(nameof(aseId), $"No existe el ASE {aseId} para el escritor de slots R1 (rango válido 1..5)."),
    };

    /// <summary>
    /// Resultado del pegado values-only: cuántas filas fuente se colocaron y cuántas celdas
    /// <c>&lt;v&gt;</c> se escribieron (F = columna Total, L = especiales). Las filas fuente con
    /// valor nulo no se escriben (queda el cero del template) y se cuentan en
    /// <c>OmitidasPorNulo</c>; los slots no consumidos quedan intactos por construcción.
    /// </summary>
    internal sealed record EscrituraSlotR1Resultado(int FilasColocadas, int CeldasF, int CeldasL, int OmitidasPorNulo);

    /// <summary>
    /// Pega los valores del bloque espejo (leído por <c>LeerEspejoR1</c>, intacto) en los slots del
    /// ASE sobre un <see cref="WorkbookPart"/> ya abierto en escritura. Solo <c>&lt;v&gt;</c>.
    /// </summary>
    internal static EscrituraSlotR1Resultado Escribir(WorkbookPart workbookPart, BloqueEspejoAseInputs bloque)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);
        ArgumentNullException.ThrowIfNull(bloque);

        var (inicio, fin) = RangoDatosAse(bloque.Ase.Id);
        var hoja = ObtenerHoja(workbookPart, HojaR1);
        VerificarCabeceraBloque(hoja, workbookPart, bloque, inicio);
        var slots = LeerSlots(hoja, workbookPart, bloque.Ase.Id, inicio, fin);
        var grupos = ConstruirGruposFirma(slots);

        var fuente = bloque.Filas;
        var fronterasFuente = ClasificarFronteras(fuente);
        var colocadas = 0;
        var celdasF = 0;
        var celdasL = 0;
        var omitidas = 0;

        for (var i = 0; i < fuente.Count; i++)
        {
            var fila = fuente[i];
            var ocurrencia = ContarOcurrenciaPrevia(fuente, fronterasFuente, i) + 1;

            // Cursor INDEPENDIENTE por firma (refinamiento T1 del hallazgo A de T4): cada fila fuente
            // agota su PROPIA lista ordenada de slots, de modo que un slot-resumen intercalado antes
            // en el template (p. ej. el AFaseo del ASE2 en F121) sigue disponible para una fila que la
            // fuente emite después. El cursor global monotónico retirado aquí dejaba ese slot atrás.
            var grupo = BuscarGrupoFirma(grupos, fila, fronterasFuente[i]);
            if (grupo is null)
            {
                throw new CalculoInvalidoException(
                    CodigoError.SinSlot,
                    MensajeSinSlot(bloque.Ase.Id, fila, fronterasFuente[i], ocurrencia, celdaCandidata: null));
            }

            var slot = grupo.SiguienteLibre();
            if (slot is null)
            {
                throw new CalculoInvalidoException(
                    CodigoError.SinSlot,
                    MensajeSinSlot(bloque.Ase.Id, fila, fronterasFuente[i], ocurrencia, grupo.UltimoConsumidoF));
            }

            var total = fila.Valor("Total");
            if (total.HasValue)
            {
                EscribirValor(slot.CeldaF, hoja, total.Value, bloque.Ase.Id, slot.Fila);
                celdasF++;
            }
            else
            {
                omitidas++;
            }

            if (bloque.TieneColumnaEspeciales)
            {
                var especiales = fila.Valor("SERVICIO ESPECIALES");
                if (especiales.HasValue)
                {
                    EscribirValor(slot.CeldaL, hoja, especiales.Value, bloque.Ase.Id, slot.Fila);
                    celdasL++;
                }
                else
                {
                    omitidas++;
                }
            }

            colocadas++;
        }

        hoja.Save();
        return new EscrituraSlotR1Resultado(colocadas, celdasF, celdasL, omitidas);
    }

    /// <summary>
    /// Slot de dato del contrato: fila del template con su firma A–E y su clasificación de
    /// frontera (solo vinculante para filas-dato de empresa). Las celdas F/L son referencias A1.
    /// </summary>
    internal sealed record SlotR1(int Fila, string CeldaF, string CeldaL, string A, string B, string C, string D, string E, bool? EsAplicacion);

    /// <summary>
    /// Grupo de slots de la MISMA firma con su propio cursor ordenado (refinamiento T1 del Plan 36,
    /// hallazgo A de T4). Sustituye al cursor global-monotónico: cada firma avanza su cursor sin
    /// interferir con las demás, de forma que un slot-resumen intercalado antes en el template sigue
    /// disponible para una fila que la fuente emite después.
    /// </summary>
    private sealed class GrupoFirma
    {
        private readonly List<SlotR1> _slots = [];
        private int _cursor;

        public GrupoFirma(SlotR1 representante) => Representante = representante;

        /// <summary>Primer slot de la firma (identidad del grupo para el matcher de firma).</summary>
        public SlotR1 Representante { get; }

        /// <summary>Celda F del último slot consumido por la firma ("(ninguno)" si aún no consumió).</summary>
        public string UltimoConsumidoF => _cursor == 0 ? "(ninguno)" : _slots[_cursor - 1].CeldaF;

        public void Agregar(SlotR1 slot) => _slots.Add(slot);

        /// <summary>Próximo slot libre de la firma en orden, o <c>null</c> si la firma se agotó.</summary>
        public SlotR1? SiguienteLibre() =>
            _cursor < _slots.Count ? _slots[_cursor++] : null;
    }

    /// <summary>
    /// Verdad si la fila fuente y el slot pertenecen a la misma firma (localizador puro):
    /// AFaseo dual-firma primero (cualquiera de las dos variantes), luego A/B/D/E exactos
    /// (OrdinalIgnoreCase) + C por sinonimia + frontera coincidente en filas-dato de empresa.
    /// Cero ramas por nombre de empresa (catálogo abierto por construcción).
    /// </summary>
    internal static bool EncajaEnSlot(FilaEspejoR1 fila, bool? fronteraFuente, SlotR1 slot)
    {
        ArgumentNullException.ThrowIfNull(fila);
        ArgumentNullException.ThrowIfNull(slot);

        if (EsFirmaAFaseo(fila.A, fila.B, fila.D, fila.E))
        {
            return EsFirmaAFaseo(slot.A, slot.B, slot.D, slot.E);
        }

        if (EsFirmaAFaseo(slot.A, slot.B, slot.D, slot.E))
        {
            return false;
        }

        if (!IgualEtiqueta(fila.A, slot.A)
            || !IgualEtiqueta(fila.B, slot.B)
            || !IgualEtiqueta(fila.D, slot.D)
            || !IgualEtiqueta(fila.E, slot.E))
        {
            return false;
        }

        if (!ColumnasCCompatibles(fila.C, slot.C))
        {
            return false;
        }

        // Plan 35-T2 como parte de la llave en filas-dato de empresa: oportuno vs aplicación no se
        // mezclan. Si algún lado no clasifica (null: sin frontera por debajo), la ocurrencia en
        // orden sigue mandando y la frontera no veta (nunca veto ciego).
        if (R1FirmaInterior.EsDatoEmpresa(fila)
            && fronteraFuente.HasValue
            && slot.EsAplicacion.HasValue
            && fronteraFuente.Value != slot.EsAplicacion.Value)
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// Firma AFaseo dual (R-T-2): variante fuente <c>A='AFaseo' B='Total'</c> o variante predicado
    /// <c>D='AFASEO' E='Total'</c>. Ambas resuelven al mismo slot F0 del ASE.
    /// </summary>
    internal static bool EsFirmaAFaseo(string a, string b, string d, string e) =>
        (IgualEtiqueta(a, "AFaseo") && IgualEtiqueta(b, "Total"))
        || (IgualEtiqueta(d, "AFASEO") && IgualEtiqueta(e, "Total"));

    private static bool IgualEtiqueta(string? x, string? y) =>
        string.Equals(x?.Trim(), y?.Trim(), StringComparison.OrdinalIgnoreCase);

    private static bool ColumnasCCompatibles(string? fuente, string? slot)
    {
        if (string.IsNullOrWhiteSpace(fuente) && string.IsNullOrWhiteSpace(slot))
        {
            return true;
        }

        if (string.IsNullOrWhiteSpace(fuente) || string.IsNullOrWhiteSpace(slot))
        {
            return false;
        }

        return SinonimosEmpresaR1.SonMismaEmpresa(fuente, slot);
    }

    /// <summary>
    /// Agrupa los slots del bloque por FIRMA: cada grupo conserva su propia lista ordenada de slots
    /// (en el orden físico del template) y su propio cursor. Un slot se suma al primer grupo cuyo
    /// representante es la MISMA firma (mismo A/B/D/E + C por sinonimia + AFaseo dual); si no encaja
    /// en ninguno, abre un grupo nuevo. Refinamiento T1 del Plan 36 (hallazgo A de T4): el cursor
    /// deja de ser global-monotónico y pasa a ser independiente por firma.
    /// </summary>
    private static List<GrupoFirma> ConstruirGruposFirma(IReadOnlyList<SlotR1> slots)
    {
        var grupos = new List<GrupoFirma>();
        foreach (var slot in slots)
        {
            var firmaSlot = FilaDesdeSlot(slot);
            var grupo = grupos.Find(g => EncajaEnSlot(firmaSlot, null, g.Representante));
            if (grupo is null)
            {
                grupo = new GrupoFirma(slot);
                grupos.Add(grupo);
            }

            grupo.Agregar(slot);
        }

        return grupos;
    }

    /// <summary>
    /// Resuelve el grupo de firma de una fila fuente. Primero exige coincidencia de firma + frontera
    /// (Plan 35-T2); si ninguna la satisface, reintenta con la firma sola: cuando la clasificación de
    /// frontera del lado fuente degenera (p. ej. la fuente no trae la fila <c>Aplicacion</c> y toda
    /// fila-dato de empresa queda "Oportuno"), el orden dentro de la firma sigue resolviendo la
    /// sección. Devuelve <c>null</c> solo si NINGUNA firma del contrato es compatible → fail-fast.
    /// </summary>
    private static GrupoFirma? BuscarGrupoFirma(List<GrupoFirma> grupos, FilaEspejoR1 fila, bool? fronteraFuente)
    {
        var conFrontera = grupos.Find(g => EncajaEnSlot(fila, fronteraFuente, g.Representante));
        return conFrontera ?? grupos.Find(g => EncajaEnSlot(fila, null, g.Representante));
    }

    private static FilaEspejoR1 FilaDesdeSlot(SlotR1 slot) => new()
    {
        A = slot.A,
        B = slot.B,
        C = slot.C,
        D = slot.D,
        E = slot.E,
    };

    private static string MensajeSinSlot(int aseId, FilaEspejoR1 fila, bool? frontera, int ocurrencia, string? celdaCandidata)
    {
        var razon = celdaCandidata is null
            ? "sin ningún slot con firma compatible en el contrato T0"
            : $"sin slot libre en el contrato T0 (celda candidata: último slot consumido {celdaCandidata})";
        return $"ERR-SIN-SLOT: ASE {aseId}: la fuente R1 (Recaudoporcomponente) trae la fila " +
            $"A|B|C|D|E='{fila.Firma}' clasificada '{TipoFrontera(frontera)}' (ocurrencia #{ocurrencia} " +
            $"de esa firma) {razon} del ASE {aseId}. Nunca remap/insert: revise el contrato T0 o la fuente.";
    }

    private static int ContarOcurrenciaPrevia(IReadOnlyList<FilaEspejoR1> fuente, IReadOnlyList<bool?> fronteras, int indice)
    {
        var conteo = 0;
        for (var k = 0; k < indice; k++)
        {
            if (EncajaEnSlot(fuente[indice], fronteras[indice], SlotDesdeFila(fuente[k], fronteras[k])))
            {
                conteo++;
            }
        }

        return conteo;
    }

    private static SlotR1 SlotDesdeFila(FilaEspejoR1 fila, bool? frontera) =>
        new(0, string.Empty, string.Empty, fila.A, fila.B, fila.C, fila.D, fila.E, frontera);

    private static string TipoFrontera(bool? frontera) => frontera switch
    {
        true => "Aplicación",
        false => "Oportuno",
        null => "sin-frontera",
    };

    private static IReadOnlyList<bool?> ClasificarFronteras(IReadOnlyList<FilaEspejoR1> filas)
    {
        var resultado = new bool?[filas.Count];
        for (var i = 0; i < filas.Count; i++)
        {
            resultado[i] = R1FirmaInterior.EsDatoEmpresa(filas[i])
                ? R1FirmaInterior.EsAplicacionPorFrontera(filas, i)
                : null;
        }

        return resultado;
    }

    private static List<SlotR1> LeerSlots(Worksheet hoja, WorkbookPart workbookPart, int aseId, int inicio, int fin)
    {
        var filasPorNumero = hoja.Elements<SheetData>().FirstOrDefault()
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{HojaR1}' no tiene SheetData para el escritor de slots R1 (ASE {aseId}).");
        var filas = filasPorNumero.Elements<Row>()
            .Where(r => r.RowIndex is not null && r.RowIndex.Value >= inicio && r.RowIndex.Value <= fin)
            .OrderBy(r => r.RowIndex!.Value)
            .ToList();

        var firmas = new List<FilaEspejoR1>();
        var celdas = new List<(string CeldaF, string CeldaL)>();
        foreach (var fila in filas)
        {
            var numero = (int)fila.RowIndex!.Value;
            firmas.Add(new FilaEspejoR1
            {
                A = TextoCelda(workbookPart, fila, $"A{numero}"),
                B = TextoCelda(workbookPart, fila, $"B{numero}"),
                C = TextoCelda(workbookPart, fila, $"C{numero}"),
                D = TextoCelda(workbookPart, fila, $"D{numero}"),
                E = TextoCelda(workbookPart, fila, $"E{numero}"),
            });
            celdas.Add(($"F{numero}", $"L{numero}"));
        }

        var fronteras = ClasificarFronteras(firmas);
        var slots = new List<SlotR1>(firmas.Count);
        for (var i = 0; i < firmas.Count; i++)
        {
            var numero = (int)filas[i].RowIndex!.Value;
            slots.Add(new SlotR1(
                numero, celdas[i].CeldaF, celdas[i].CeldaL,
                firmas[i].A, firmas[i].B, firmas[i].C, firmas[i].D, firmas[i].E,
                fronteras[i]));
        }

        if (slots.Count != fin - inicio + 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El bloque R1 del ASE {aseId} trae {slots.Count} filas físicas en [{inicio},{fin}] (se esperaban {fin - inicio + 1} filas contiguas del contrato T0).");
        }

        return slots;
    }

    /// <summary>
    /// Pega el valor del slot como <c>&lt;v&gt;</c> numérico. Contrato de redondeo (W1 del review de
    /// T1, decidido aquí): el slot R1 conserva el DECIMAL EXACTO de la fuente — NO se redondea.
    ///
    /// Por qué: el §2.1 del plan dice "número redondeado según contrato", y el contrato que manda
    /// (T0 congelado desde disco + el comportamiento histórico del writer por encabezado) es
    /// "mismo valor de la fuente": los slots F/L guardan importes con decimales (p. ej.
    /// <c>F649 = 13359789539.23</c> en el template) y el gate los compara contra el dato de la
    /// fuente. El redondeo a ENTERO es exclusivo del <c>DetRetri</c> (columna D, Plan 29 Unidad D,
    /// <c>DetRetriRounder</c>) y vive allí, no aquí. Redondear en el slot introduciría una
    /// divergencia silenciosa contra la fuente dentro de la tolerancia ±0.5 y rompería la
    /// igualdad exacta que el contrato T0/T1 aserta. Guard anti-fórmula: celda destino con
    /// <c>&lt;f&gt;</c> → <c>ERR-PLANTILLA</c> (nunca se sobrescribe una fórmula).
    /// </summary>
    private static void EscribirValor(string celda, Worksheet hoja, decimal valor, int aseId, int filaSlot)
    {
        var cell = hoja.Descendants<Cell>().FirstOrDefault(c =>
                string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La celda slot '{HojaR1}!{celda}' (fila {filaSlot} del contrato T0, ASE {aseId}) no existe en la plantilla. Fail-fast, nunca se inventa la celda.");

        if (cell.CellFormula is not null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La celda slot '{HojaR1}!{celda}' (fila {filaSlot} del contrato T0, ASE {aseId}) es fórmula en la plantilla. El escritor values-only nunca sobrescribe fórmulas.");
        }

        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture));
    }

    private static void VerificarCabeceraBloque(Worksheet hoja, WorkbookPart workbookPart, BloqueEspejoAseInputs bloque, int inicio)
    {
        // La cabecera del bloque vive dos filas arriba del primer slot (R4/R108/R287/R448/R601,
        // congeladas por T0). Se exige identidad con el catálogo vigente (CarpetasAse vía
        // AseFactory) usando el matcher de sinonimia del Plan 34 como ÚNICA normalización de
        // rótulos de empresa (S1 del review de T1: cero duplicado local de la normalización
        // mayúsculas/sin-diacríticos de SinonimosEmpresaR1): un workbook con el bloque equivocado
        // hace fail-fast antes de pegar ningún valor.
        var filaCabecera = hoja.Elements<SheetData>().FirstOrDefault()?
            .Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == inicio - 2);
        if (filaCabecera is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El bloque R1 del ASE {bloque.Ase.Id} no trae la fila de cabecera {inicio - 2} del contrato T0.");
        }

        var rotulo = TextoCelda(workbookPart, filaCabecera, $"B{inicio - 2}");
        var esperado = AseFactory.DesdeId(bloque.Ase.Id).NombreCompleto;
        if (!SinonimosEmpresaR1.SonMismaEmpresa(rotulo, esperado))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La cabecera del bloque R1 ('{HojaR1}!B{inicio - 2}' = '{rotulo}') no corresponde al ASE {bloque.Ase.Id} ('{esperado}'). Fail-fast antes de escribir.");
        }
    }

    private static Worksheet ObtenerHoja(WorkbookPart workbookPart, string nombreHoja)
    {
        var hoja = workbookPart.Workbook?.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"El workbook no trae la hoja '{nombreHoja}' para el escritor de slots R1.");
        var parte = (WorksheetPart)workbookPart.GetPartById(hoja.Id!);
        return parte.Worksheet
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{nombreHoja}' no tiene WorksheetPart válido.");
    }

    private static string TextoCelda(WorkbookPart workbookPart, Row fila, string referencia)
    {
        var celda = fila.Elements<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        if (celda is null)
        {
            return string.Empty;
        }

        if (celda.DataType is not null && celda.DataType.Value == CellValues.SharedString
            && celda.CellValue is not null
            && int.TryParse(celda.CellValue.Text, out var indice))
        {
            var tabla = workbookPart.SharedStringTablePart?.SharedStringTable;
            var item = tabla?.Elements<SharedStringItem>().ElementAtOrDefault(indice);
            return item?.InnerText ?? string.Empty;
        }

        if (celda.DataType is not null && celda.DataType.Value == CellValues.InlineString)
        {
            return celda.InlineString?.InnerText ?? string.Empty;
        }

        return celda.CellValue?.InnerText ?? string.Empty;
    }
}
