using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Serilog;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 36 (T1, R-B-1/R-B-2/R-B-4 — D-A/D-B/D-F) + Plan 41 (T1/Enmienda E-1) + Plan 44 (F1/Enmienda
/// E-44.1): escritor values-only MULTICOLUMNA del interior de <c>Reporte Componentes R1</c> sobre la
/// plantilla estandarizada única.
///
/// Doctrina slot: la app NUNCA inserta/borra filas y NUNCA escribe <c>&lt;f&gt;</c>; solo pega
/// VALORES (<c>&lt;v&gt;</c>) en el slot de datos que corresponde por firma + SECCIÓN + sinonimia;
/// lo no entregado queda en el cero del template; la fuente que trae fila sin slot hace fail-fast
/// <c>ERR-SIN-SLOT</c> (nunca remap silencioso, nunca insert, nunca colocación por posición).
///
/// Resolución (arsenal reinvertido, cero lógica nueva de negocio):
/// <list type="bullet">
/// <item>firma como localizador (<see cref="R1FirmaInterior"/> — mismos predicados que hoy);</item>
/// <item>SECCIÓN por frontera inmediatamente ARRIBA (<see cref="SeccionEspejoR1.PorFronteraArriba"/>,
/// Plan 41/T1: <c>Mes</c>→Oportuno, <c>Aplicacion</c>→Aplicación, sin frontera→1.ª), exigida en
/// filas-dato de empresa (<see cref="R1FirmaInterior.EsDatoEmpresa(FilaEspejoR1)"/>). Es la
/// resolución ÚNICA para esas filas, sin fallback (Enmienda E-1, E1-1);</item>
/// <item>empresa por sinonimia (<see cref="SinonimosEmpresaR1.SonMismaEmpresa"/>, Plan 34 intacto,
/// catálogo abierto — cero ramas por nombre de empresa);</item>
/// <item>AFaseo dual-firma (R-T-2): la fuente en variante <c>A='AFaseo' B='Total'</c> o en variante
/// <c>D='AFASEO' E='Total'</c> resuelve al mismo slot F0 del ASE.</item>
/// </list>
///
/// Plan 44 (F1/E-44.1) + Plan 45 (H1/S2): el LLENADO MULTICOLUMNA (F..Z por encabezado) está
/// implementado; la SECCIÓN-CONTRATO para las filas sin empresa en C (P2/E-44.1) la fija el contrato
/// PERIOD-FREE por ASE (<see cref="ContratoSeccionPeriodFreeR1"/>, Plan 45): alineación monotónica por
/// tipo fuente↔oficial que absorbe las 3 formas probadas (julio-2 invertida de ASE5, agosto con 3
/// <c>Mes</c>, Q1 sin <c>Aplicacion</c>) sin literales ni ramas por período/empresa (D-H). Las filas
/// sin slot de la lista T0 caen a CERO DOCUMENTADO (<see cref="CerosDocumentadosR1"/>, no es error);
/// fuera de la lista es <c>ERR-SIN-SLOT</c> fail-fast.
///
/// Contrato de columnas (Plan 44 F1, §2.3 del plan + T0 §1/§7): la plantilla resuelve las columnas
/// destino por ENCABEZADO del template (F..Z; la PRIMERA ocurrencia de un encabezado repetido manda:
/// el `Intereses Cte Financiac. Aseo` de la fuente va a **T**, y **Y** queda intacta). El
/// `Rec.Intereses` de la fuente (col U) va a **Z** de la plantilla. Las columnas `Exonera-*` (U/V/W/X)
/// del template NO tienen contraparte en la fuente y quedan INTACTAS por construcción.
///
/// Contrato de valores (W1 del review de T1): el slot recibe el DECIMAL EXACTO de la fuente (sin
/// redondeo); el redondeo a entero es exclusivo del <c>DetRetri</c> (Plan 29 Unidad D). Ver la
/// justificación completa en <see cref="EscribirValor"/>.
///
/// Emparejamiento por (firma + SECCIÓN) con cursor INDEPENDIENTE por grupo: en la plantilla
/// estandarizada los slots-resumen (AFaseo, Mes, …) están INTERCALADOS dentro del bloque y la fuente
/// puede emitirlos en un orden distinto — p. ej. el AFaseo del ASE2 vive en F121 (en medio del
/// bloque) mientras la fuente lo emite DESPUÉS de consumir filas posteriores del template. Un cursor
/// GLOBAL monotónico no puede retroceder y deja ese AFaseo sin slot (falso <c>ERR-SIN-SLOT</c>). Por
/// eso cada fila fuente consume el próximo slot libre DENTRO DE SU PROPIA (firma, sección). Los
/// slots no consumidos quedan intactos en cero. Guard pre-write (REGLA, no excepción): la celda
/// destino debe existir y NO traer <c>&lt;f&gt;</c>; si trae fórmula → <c>ERR-PLANTILLA</c> con hoja
/// + celda + slot. Cero <c>InsertRow</c>/<c>DeleteRow</c>, cero <c>new CellFormula</c>, cero Reanclar
/// en este path (el grep de cierre T1 lo demuestra).
///
/// El writer NO recibe ningún período (D-H) y no nombra ninguna empresa (catálogo abierto).
/// <see cref="CompositorInteriorR1"/> queda como referencia de firma, no se usa aquí (D-C).
/// </summary>
internal static class EscritorValoresSlotR1
{
    public const string HojaR1 = "Reporte Componentes R1";

    /// <summary>
    /// Columnas del bloque R1 de la plantilla con encabezado (F..Z), en orden físico. El writer
    /// resuelve el mapeo encabezado→letra una vez por hoja (Plan 44 F1).
    /// </summary>
    private static readonly string[] ColumnasComponentes =
    [
        "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
    ];

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
    /// Resultado del pegado values-only: cuántas filas fuente se colocaron, cuántas celdas
    /// <c>&lt;v&gt;</c> se escribieron en las columnas F (Total) y L (especiales) — las del contrato
    /// histórico —, cuántas filas quedaron en el cero del template por valor nulo y cuántas cayeron en
    /// <b>cero documentado</b> (Plan 45 §2.2: fila sin slot de la lista T0; no se escribe). Los slots
    /// no consumidos quedan intactos por construcción. <see cref="EncabezadosSinDestino"/> recoge los
    /// encabezados de la fuente SIN columna destino en el template (H1.4; nunca descarte silencioso).
    /// </summary>
    internal sealed record EscrituraSlotR1Resultado(
        int FilasColocadas,
        int CeldasF,
        int CeldasL,
        int OmitidasPorNulo,
        int CerosDocumentados,
        IReadOnlyList<string> EncabezadosSinDestino);

    /// <summary>
    /// Pega los valores del bloque espejo (leído por <c>LeerEspejoR1</c>, intacto) en los slots del
    /// ASE sobre un <see cref="WorkbookPart"/> ya abierto en escritura. Solo <c>&lt;v&gt;</c>.
    /// </summary>
    internal static EscrituraSlotR1Resultado Escribir(WorkbookPart workbookPart, BloqueEspejoAseInputs bloque)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);
        ArgumentNullException.ThrowIfNull(bloque);

        var aseId = bloque.Ase.Id;
        var (inicio, fin) = RangoDatosAse(aseId);
        var hoja = ObtenerHoja(workbookPart, HojaR1);
        VerificarCabeceraBloque(hoja, workbookPart, bloque, inicio);
        var mapaColumnas = LeerMapaColumnasDestino(hoja, workbookPart, aseId, inicio);
        var fuente = bloque.Filas;

        // Plan 45 (§2.2) + NEUTRALIDAD Q1 (decisión del Ingeniero 2026-10-16, opción 1): la resolución
        // es la SECCIÓN-CONTRATO PERIOD-FREE (Plan 45) con CERO DOCUMENTADO. Si el bloque ENTERO no es
        // colocable con ella (la forma del período no está cubierta por el contrato — p. ej. Q1-ASE2
        // trae el canal D=5 en la 1.ª, que el oficial no declara), se reproduce el comportamiento
        // HISTÓRICO del Plan 41 (veto de sección SOLO en filas-dato de empresa; cursor por firma) para
        // ese bloque: así Q1 conserva EXACTAMENTE su primer fail-fast (nombra NUEVO ESQUEMA, no D=5) y
        // ninguna fila "gana" un fail-fast que el comportamiento previo no daba. Sin ramas por
        // período/empresa: la decisión es por FORMA colocable, geometría pura (D-H).
        var slots = LeerSlots(hoja, workbookPart, aseId, inicio, fin, porContrato: true);
        if (!TryPlanificar(fuente, slots, porContrato: true, out var plan, out var ceros, out _))
        {
            slots = LeerSlots(hoja, workbookPart, aseId, inicio, fin, porContrato: false);
            if (!TryPlanificar(fuente, slots, porContrato: false, out plan, out ceros, out var fallo))
            {
                throw new CalculoInvalidoException(
                    CodigoError.SinSlot,
                    MensajeSinSlot(aseId, fallo.Fila, fallo.Seccion, fallo.Ocurrencia, fallo.CeldaCandidata));
            }
        }

        var colocadas = 0;
        var celdasF = 0;
        var celdasL = 0;
        var omitidas = 0;
        var encabezadosSinDestino = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (fila, slot) in plan)
        {
            EscribirComponentes(hoja, mapaColumnas, slot, fila, aseId, ref celdasF, ref celdasL, encabezadosSinDestino);
            colocadas++;
            if (fila.Valor("Total") is null)
            {
                omitidas++;
            }
        }

        hoja.Save();
        return new EscrituraSlotR1Resultado(colocadas, celdasF, celdasL, omitidas, ceros, encabezadosSinDestino.ToList());
    }

    /// <summary>Punto de fail-fast planificado (alimenta el mensaje <c>ERR-SIN-SLOT</c>).</summary>
    private readonly record struct FalloSinSlot(FilaEspejoR1 Fila, SeccionFronteraArriba? Seccion, int Ocurrencia, string? CeldaCandidata);

    /// <summary>
    /// Planifica (SIN escribir) el pegado del bloque completo: <paramref name="porContrato"/> <c>true</c>
    /// = sección-contrato PERIOD-FREE + cero documentado (Plan 45); <c>false</c> = comportamiento
    /// histórico Plan 41 (veto de sección solo en filas-dato de empresa, cursor por firma). Devuelve las
    /// colocaciones (fila → slot) en orden, el nº de ceros documentados y, si no es colocable, el punto
    /// de fail-fast. Un intento fallido no escribe nada (la escritura ocurre tras elegir el plan).
    /// </summary>
    private static bool TryPlanificar(
        IReadOnlyList<FilaEspejoR1> fuente,
        IReadOnlyList<SlotR1> slots,
        bool porContrato,
        out List<(FilaEspejoR1 Fila, SlotR1 Slot)> plan,
        out int cerosDocumentados,
        out FalloSinSlot fallo)
    {
        var grupos = ConstruirGruposSeccion(slots);
        var secciones = porContrato ? ClasificarSeccionesContrato(fuente, slots) : ClasificarSeccionesBaseline(fuente);
        plan = [];
        cerosDocumentados = 0;
        fallo = default;

        for (var i = 0; i < fuente.Count; i++)
        {
            var fila = fuente[i];
            var seccion = secciones[i];
            var grupo = BuscarGrupo(grupos, fila, seccion);
            var slot = grupo?.SiguienteLibre();

            if (slot is null)
            {
                if (porContrato && CerosDocumentadosR1.EsCeroDocumentado(fila, seccion))
                {
                    cerosDocumentados++;
                    continue;
                }

                fallo = new FalloSinSlot(fila, seccion, ContarOcurrenciaPrevia(fuente, secciones, i) + 1, grupo?.UltimoConsumidoF);
                return false;
            }

            plan.Add((fila, slot));
        }

        return true;
    }

    /// <summary>
    /// Slot de dato del contrato: fila del template con su firma A–E y su SECCIÓN-CONTRATO (frontera
    /// inmediatamente ARRIBA para las filas de datos; la sección que ABREN para las filas-frontera;
    /// <c>null</c> para AFaseo). Las celdas F/L son referencias A1.
    /// </summary>
    internal sealed record SlotR1(int Fila, string CeldaF, string CeldaL, string A, string B, string C, string D, string E, SeccionFronteraArriba? Seccion);

    /// <summary>
    /// Grupo de slots de la MISMA firma y SECCIÓN-CONTRATO (Plan 41/T1 + Plan 45 §2.2) con su propio
    /// cursor ordenado. Cada (firma, sección-contrato) avanza su cursor sin interferir con los demás:
    /// un slot-resumen intercalado antes en el template sigue disponible para una fila que la fuente
    /// emite después. Las filas-frontera (<c>Mes/Total</c>, <c>Aplicacion…/Total</c>) llevan la sección
    /// que ABREN; solo AFaseo queda <c>null</c> (grupo por firma, comportamiento histórico). Plan 41
    /// (Enmienda E-1, E1-1): NO existe ningún cursor de respaldo — este es el ÚNICO mecanismo.
    /// </summary>
    private sealed class GrupoSeccion
    {
        private readonly List<SlotR1> _slots = [];
        private int _cursor;

        public GrupoSeccion(SlotR1 representante) => Representante = representante;

        /// <summary>Primer slot del grupo (identidad del grupo para el matcher firma+sección).</summary>
        public SlotR1 Representante { get; }

        /// <summary>Celda F del último slot consumido por el grupo ("(ninguno)" si aún no consumió).</summary>
        public string UltimoConsumidoF => _cursor == 0 ? "(ninguno)" : _slots[_cursor - 1].CeldaF;

        public void Agregar(SlotR1 slot) => _slots.Add(slot);

        /// <summary>Próximo slot libre del grupo en orden, o <c>null</c> si el grupo se agotó.</summary>
        public SlotR1? SiguienteLibre() => _cursor < _slots.Count ? _slots[_cursor++] : null;

        /// <summary>Verdad si el grupo aún tiene un slot libre (sin consumirlo).</summary>
        public bool TieneLibre => _cursor < _slots.Count;
    }

    /// <summary>
    /// Verdad si la fila fuente y el slot pertenecen a la misma firma + sección-contrato (localizador
    /// puro): AFaseo dual-firma primero (cualquiera de las dos variantes), luego A/B/D/E exactos
    /// (OrdinalIgnoreCase) + C por sinonimia; y —cuando AMBOS lados traen sección— la SECCIÓN debe
    /// coincidir. Solo AFaseo viaja sin sección (se localiza por firma). Cero ramas por nombre de
    /// empresa (catálogo abierto por construcción).
    /// </summary>
    internal static bool EncajaEnSlot(FilaEspejoR1 fila, SeccionFronteraArriba? seccionFuente, SlotR1 slot)
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

        // Plan 41-T1 (D-A) + Plan 45 (§2.2) como parte de la llave: Oportuno vs Aplicación y las
        // ocurrencias NO se mezclan. Si AMBOS lados traen sección (filas de datos/componentes) deben
        // coincidir; si algún lado no participa (null: fila-frontera/AFaseo — se localizan por firma),
        // la sección no veta.
        if (seccionFuente.HasValue
            && slot.Seccion.HasValue
            && seccionFuente.Value != slot.Seccion.Value)
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
    /// Agrupa los slots del bloque por FIRMA + SECCIÓN (Plan 41/T1, D-A): cada grupo conserva su
    /// propia lista ordenada de slots (en el orden físico del template) y su propio cursor. Un slot
    /// se suma al primer grupo cuyo representante coincide (misma firma A/B/D/E + C por sinonimia +
    /// AFaseo dual + misma sección en filas de datos); si no encaja en ninguno, abre un grupo nuevo.
    /// </summary>
    private static List<GrupoSeccion> ConstruirGruposSeccion(IReadOnlyList<SlotR1> slots)
    {
        var grupos = new List<GrupoSeccion>();
        foreach (var slot in slots)
        {
            var firmaSlot = FilaDesdeSlot(slot);
            var grupo = grupos.Find(g => EncajaEnSlot(firmaSlot, slot.Seccion, g.Representante));
            if (grupo is null)
            {
                grupo = new GrupoSeccion(slot);
                grupos.Add(grupo);
            }

            grupo.Agregar(slot);
        }

        return grupos;
    }

    private static FilaEspejoR1 FilaDesdeSlot(SlotR1 slot) => new()
    {
        A = slot.A,
        B = slot.B,
        C = slot.C,
        D = slot.D,
        E = slot.E,
    };

    /// <summary>
    /// Primer grupo (en orden físico del template) que encaja con la fila por firma + (<paramref name="seccion"/>
    /// o sin sección si es <c>null</c>) Y aún tiene un slot libre. Con <paramref name="seccion"/> no nulo
    /// es la resolución por SECCIÓN-CONTRATO (Plan 45 §2.2); con <c>null</c> es el POOL HISTÓRICO
    /// (firma sola, Plan 41) que preserva la neutralidad Q1 (nunca un fail-fast antes que el previo).
    /// </summary>
    private static GrupoSeccion? BuscarGrupo(IReadOnlyList<GrupoSeccion> grupos, FilaEspejoR1 fila, SeccionFronteraArriba? seccion)
    {
        foreach (var grupo in grupos)
        {
            if (grupo.TieneLibre && EncajaEnSlot(fila, seccion, grupo.Representante))
            {
                return grupo;
            }
        }

        return null;
    }

    /// <summary>
    /// Mensaje de <c>ERR-SIN-SLOT</c> (Plan 41 E1-3 + Plan 44 E-44.1): nombra la FILA (firma
    /// A|B|C|D|E + ocurrencia), su RÓTULO DE EMPRESA (columna C), el CANAL (columna D) y el CONCEPTO
    /// (columna E) de la fila, y la SECCIÓN ESPERADA (la resuelta por su frontera inmediatamente
    /// arriba y mapeada por el contrato) — información accionable, nunca un fallo mudo.
    /// </summary>
    private static string MensajeSinSlot(int aseId, FilaEspejoR1 fila, SeccionFronteraArriba? seccion, int ocurrencia, string? celdaCandidata)
    {
        var razon = celdaCandidata is null
            ? "sin ningún slot con firma + sección compatibles en el contrato T0"
            : $"sin slot libre en el contrato T0 (celda candidata: último slot consumido {celdaCandidata})";
        return $"ERR-SIN-SLOT: ASE {aseId}: la fuente R1 (Recaudoporcomponente) trae la fila " +
            $"A|B|C|D|E='{fila.Firma}' con rótulo de empresa '{fila.C}', canal '{fila.D}' y concepto '{fila.E}' " +
            $"en la sección esperada '{TipoSeccion(seccion)}' (ocurrencia #{ocurrencia} de esa firma) {razon} del ASE {aseId}. " +
            $"Nunca remap/insert ni colocación por posición: revise el contrato T0 (slot por sección) o la fuente.";
    }

    private static int ContarOcurrenciaPrevia(IReadOnlyList<FilaEspejoR1> fuente, IReadOnlyList<SeccionFronteraArriba?> secciones, int indice)
    {
        var conteo = 0;
        for (var k = 0; k < indice; k++)
        {
            if (EncajaEnSlot(fuente[indice], secciones[indice], SlotDesdeFila(fuente[k], secciones[k])))
            {
                conteo++;
            }
        }

        return conteo;
    }

    private static SlotR1 SlotDesdeFila(FilaEspejoR1 fila, SeccionFronteraArriba? seccion) =>
        new(0, string.Empty, string.Empty, fila.A, fila.B, fila.C, fila.D, fila.E, seccion);

    private static string TipoSeccion(SeccionFronteraArriba? seccion) => seccion switch
    {
        null => "sin-sección",
        { Tipo: TipoSeccionEspejoR1.Primera } => "1.ª sección",
        { Tipo: TipoSeccionEspejoR1.Oportuno } => $"Oportuno #{seccion.Value.Ordinal}",
        { Tipo: TipoSeccionEspejoR1.Aplicacion } => $"Aplicación #{seccion.Value.Ordinal}",
        _ => $"desconocida #{seccion.Value.Ordinal}",
    };

    /// <summary>
    /// Sección CRUDA de cada fila (Plan 41/T1 + Plan 45 §2.2): la frontera inmediatamente ARRIBA para
    /// las filas de datos/componentes; para las filas-frontera (<c>Mes/Total</c>, <c>Aplicacion…/Total</c>)
    /// la sección que ABREN (la de su primera fila hija) — así la frontera viaja con la sección a la que
    /// pertenece por el contrato; AFaseo queda <c>null</c> (se localiza por firma). Es la sección NATIVA
    /// de la hoja (del oficial para los slots; de la fuente para las filas fuente antes de mapear).
    /// </summary>
    private static IReadOnlyList<SeccionFronteraArriba?> ClasificarSeccionesRaw(IReadOnlyList<FilaEspejoR1> filas)
    {
        var resultado = new SeccionFronteraArriba?[filas.Count];
        for (var i = 0; i < filas.Count; i++)
        {
            var fila = filas[i];
            if (EsFirmaAFaseo(fila.A, fila.B, fila.D, fila.E))
            {
                resultado[i] = null;
            }
            else if (fila.EsMesTotal || fila.EsAplicacionTotal)
            {
                resultado[i] = SeccionQueAbre(filas, i);
            }
            else
            {
                resultado[i] = SeccionEspejoR1.PorFronteraArriba(filas, i);
            }
        }

        return resultado;
    }

    /// <summary>
    /// Sección que ABRE la fila-frontera <paramref name="indice"/>: la sección de su primera fila hija
    /// (la próxima hacia abajo que participa de la alineación por sección). Para la frontera
    /// <c>Mes</c> la hija abre <c>Oportuno</c>; para <c>Aplicacion</c>, <c>Aplicación</c>.
    /// </summary>
    private static SeccionFronteraArriba SeccionQueAbre(IReadOnlyList<FilaEspejoR1> filas, int indice)
    {
        for (var j = indice + 1; j < filas.Count; j++)
        {
            if (ContratoSeccionPeriodFreeR1.ParticipaDeSeccionConformada(filas[j]))
            {
                return SeccionEspejoR1.PorFronteraArriba(filas, j);
            }
        }

        // Frontera sin hija (cierre vacío): su propia sección por frontera arriba.
        return SeccionEspejoR1.PorFronteraArriba(filas, indice);
    }

    /// <summary>
    /// Sección del comportamiento HISTÓRICO (Plan 41): la frontera inmediatamente ARRIBA SOLO en las
    /// filas-dato de empresa (<see cref="R1FirmaInterior.EsDatoEmpresa"/>); <c>null</c> para el resto
    /// (cursor por firma). Es la reproducción byte-fiel del comportamiento previo a la regla sin-C —
    /// la usa el plan de NEUTRALIDAD Q1 cuando el bloque no es colocable por el contrato.
    /// </summary>
    private static IReadOnlyList<SeccionFronteraArriba?> ClasificarSeccionesBaseline(IReadOnlyList<FilaEspejoR1> filas)
    {
        var resultado = new SeccionFronteraArriba?[filas.Count];
        for (var i = 0; i < filas.Count; i++)
        {
            resultado[i] = R1FirmaInterior.EsDatoEmpresa(filas[i])
                ? SeccionEspejoR1.PorFronteraArriba(filas, i)
                : null;
        }

        return resultado;
    }

    /// <summary>
    /// Sección-CONTRATO de cada fila FUENTE (Plan 45 §2.1/§2.2): la sección cruda de la fuente mapeada
    /// a la sección del OFICIAL por el contrato PERIOD-FREE (<see cref="ContratoSeccionPeriodFreeR1.Alinear"/>,
    /// alineación monotónica por tipo). Una sección fuente sin contraparte del mismo tipo conserva su
    /// sección cruda (sus filas fallarán como ERR-SIN-SLOT o caerán a cero documentado — nunca remap por
    /// posición). Las fronteras/AFaseo quedan en <c>null</c> (se localizan por firma).
    /// </summary>
    private static IReadOnlyList<SeccionFronteraArriba?> ClasificarSeccionesContrato(
        IReadOnlyList<FilaEspejoR1> fuente,
        IReadOnlyList<SlotR1> slots)
    {
        var crudas = ClasificarSeccionesRaw(fuente);
        var seccionesOficial = DistintasOrdenadas(slots.Select(s => s.Seccion));
        var mapa = ContratoSeccionPeriodFreeR1.Alinear(DistintasOrdenadas(crudas), seccionesOficial);

        var resultado = new SeccionFronteraArriba?[fuente.Count];
        for (var i = 0; i < fuente.Count; i++)
        {
            resultado[i] = crudas[i] is { } cruda && mapa.TryGetValue(cruda, out var destino)
                ? destino
                : crudas[i];
        }

        return resultado;
    }

    /// <summary>Secciones distintas en el orden en que aparecen (ignora los <c>null</c>).</summary>
    private static List<SeccionFronteraArriba> DistintasOrdenadas(IEnumerable<SeccionFronteraArriba?> secciones)
    {
        var lista = new List<SeccionFronteraArriba>();
        SeccionFronteraArriba? previa = null;
        foreach (var seccion in secciones)
        {
            if (seccion is null)
            {
                continue;
            }

            if (previa is null || previa.Value != seccion.Value)
            {
                lista.Add(seccion.Value);
            }

            previa = seccion;
        }

        return lista;
    }

    private static List<SlotR1> LeerSlots(Worksheet hoja, WorkbookPart workbookPart, int aseId, int inicio, int fin, bool porContrato)
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

        var secciones = porContrato ? ClasificarSeccionesRaw(firmas) : ClasificarSeccionesBaseline(firmas);
        var slots = new List<SlotR1>(firmas.Count);
        for (var i = 0; i < firmas.Count; i++)
        {
            var numero = (int)filas[i].RowIndex!.Value;
            slots.Add(new SlotR1(
                numero, celdas[i].CeldaF, celdas[i].CeldaL,
                firmas[i].A, firmas[i].B, firmas[i].C, firmas[i].D, firmas[i].E,
                secciones[i]));
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
    /// Plan 44 (F1, §2.3): resuelve el mapeo ENCABEZADO→LETRA de las columnas del bloque R1 desde la
    /// fila de encabezados del template (<c>inicio - 1</c>). Un encabezado repetido conserva la
    /// PRIMERA ocurrencia (T0 §7: `Intereses Cte Financiac. Aseo` → **T**, no Y). Las columnas
    /// `Exonera-*` quedan en el mapa pero la fuente no las trae → nunca se escriben.
    /// </summary>
    private static IReadOnlyDictionary<string, string> LeerMapaColumnasDestino(Worksheet hoja, WorkbookPart workbookPart, int aseId, int inicio)
    {
        var filaHeader = hoja.Elements<SheetData>().FirstOrDefault()?
            .Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == inicio - 1)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El bloque R1 del ASE {aseId} no trae la fila de encabezados {inicio - 1} del contrato de columnas por encabezado (Plan 44 F1).");

        var mapa = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var letra in ColumnasComponentes)
        {
            var texto = TextoCelda(workbookPart, filaHeader, $"{letra}{inicio - 1}").Trim();
            if (texto.Length == 0 || mapa.ContainsKey(texto))
            {
                continue;
            }

            mapa[texto] = letra;
        }

        return mapa;
    }

    /// <summary>
    /// Plan 44 (F1): pega TODOS los componentes de la fila fuente en el slot, por ENCABEZADO
    /// (F..Z). Solo escribe los encabezados que existen en la fuente Y tienen columna destino en el
    /// template; los valores nulos se omiten (queda el cero del template). Decimal exacto.
    /// </summary>
    private static void EscribirComponentes(
        Worksheet hoja,
        IReadOnlyDictionary<string, string> mapaColumnas,
        SlotR1 slot,
        FilaEspejoR1 fila,
        int aseId,
        ref int celdasF,
        ref int celdasL,
        HashSet<string> encabezadosSinDestino)
    {
        foreach (var (encabezado, valor) in fila.ValoresPorColumna)
        {
            if (valor is null)
            {
                continue;
            }

            if (!mapaColumnas.TryGetValue(encabezado, out var columna))
            {
                // H1.4 (review F1 SPT-1): encabezado de la fuente SIN columna destino en el template.
                // Nunca descarte silencioso: se registra accionable (ASE + fila + encabezado) y se
                // recoge en el resultado. Hoy inocuo (T0 §1.2: F..U fuente mapea 1:1 a F..Z).
                encabezadosSinDestino.Add(encabezado);
                Log.Warning(
                    "{Hoja}: encabezado de fuente sin columna destino '{Encabezado}' (ASE {Ase}, fila slot {Fila}); valor descartado — revise el contrato de columnas (H1.4).",
                    HojaR1,
                    encabezado,
                    aseId,
                    slot.Fila);
                continue;
            }

            EscribirValor($"{columna}{slot.Fila}", hoja, valor.Value, aseId, slot.Fila);
            if (string.Equals(columna, "F", StringComparison.OrdinalIgnoreCase))
            {
                celdasF++;
            }
            else if (string.Equals(columna, "L", StringComparison.OrdinalIgnoreCase))
            {
                celdasL++;
            }
        }
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
