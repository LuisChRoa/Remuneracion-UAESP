using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 36 (T1, R-B-1/R-B-2/R-B-4 — D-A/D-B/D-F) + Plan 41 (T1/Enmienda E-1): escritor values-only
/// del interior de <c>Reporte Componentes R1</c> sobre la plantilla estandarizada única.
///
/// Doctrina slot: la app NUNCA inserta/borra filas y NUNCA escribe <c>&lt;f&gt;</c>; solo pega
/// VALORES (<c>&lt;v&gt;</c>) en el slot de datos que corresponde por firma + SECCIÓN + sinonimia;
/// lo no entregado queda en el cero del template; la fuente que trae empresa/fila sin slot hace
/// fail-fast <c>ERR-SIN-SLOT</c> (nunca remap silencioso, nunca insert, nunca colocación por
/// posición).
///
/// Resolución (arsenal reinvertido, cero lógica nueva de negocio):
/// <list type="bullet">
/// <item>firma como localizador (<see cref="R1FirmaInterior"/> — mismos predicados que hoy);</item>
/// <item>SECCIÓN por frontera inmediatamente ARRIBA (<see cref="SeccionEspejoR1.PorFronteraArriba"/>,
/// Plan 41/T1: <c>Mes</c>→Oportuno, <c>Aplicacion</c>→Aplicación, sin frontera→1.ª), exigida en
/// filas-dato de empresa (<see cref="R1FirmaInterior.EsDatoEmpresa(FilaEspejoR1)"/>). Es la
/// resolución ÚNICA, sin fallback (Enmienda E-1, E1-1). El predicado de frontera por DEBAJO
/// (<see cref="R1FirmaInterior.EsAplicacionPorFrontera"/>, Plan 35-T2) queda INTACTO para su otro
/// consumidor;</item>
/// <item>empresa por sinonimia (<see cref="SinonimosEmpresaR1.SonMismaEmpresa"/>, Plan 34 intacto,
/// catálogo abierto — cero ramas por nombre de empresa); el rótulo de la fila fuente se normaliza
/// por ese catálogo antes de resolver su sección (Enmienda E-1, E1-2);</item>
/// <item>AFaseo dual-firma (R-T-2): la fuente en variante <c>A='AFaseo' B='Total'</c> o en variante
/// <c>D='AFASEO' E='Total'</c> resuelve al mismo slot F0 del ASE.</item>
/// </list>
///
/// Contrato de valores (W1 del review de T1): el slot recibe el DECIMAL EXACTO de la fuente (sin
/// redondeo); el redondeo a entero es exclusivo del <c>DetRetri</c> (Plan 29 Unidad D) y no se
/// aplica aquí. Ver la justificación completa en <see cref="EscribirValor"/>.
///
/// Emparejamiento por (firma + SECCIÓN) con cursor INDEPENDIENTE por grupo: en la plantilla
/// estandarizada los slots-resumen (AFaseo, Mes, …) están INTERCALADOS dentro del bloque y la fuente
/// puede emitirlos en un orden distinto — p. ej. el AFaseo del ASE2 vive en F121 (en medio del
/// bloque) mientras la fuente lo emite DESPUÉS de consumir filas posteriores del template. Un cursor
/// GLOBAL monotónico no puede retroceder y deja ese AFaseo sin slot (falso <c>ERR-SIN-SLOT</c>). Por
/// eso cada fila fuente consume el próximo slot libre DENTRO DE SU PROPIA (firma, sección) (lista
/// ordenada de slots compatibles): un slot intercalado antes en el template puede ser consumido por
/// una fila emitida después, y cada grupo avanza su propio cursor sin interferir con los demás. Los
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
        var grupos = ConstruirGruposSeccion(slots);

        var fuente = bloque.Filas;
        var seccionesFuente = ClasificarSecciones(fuente);
        var colocadas = 0;
        var celdasF = 0;
        var celdasL = 0;
        var omitidas = 0;

        for (var i = 0; i < fuente.Count; i++)
        {
            var fila = fuente[i];
            var ocurrencia = ContarOcurrenciaPrevia(fuente, seccionesFuente, i) + 1;

            // Plan 41 (Enmienda E-1, E1-1/E1-2/E1-3): la colocación es SIEMPRE por (firma + SECCIÓN
            // RESUELTA), sin ningún fallback por cursor físico (ni global, ni por firma, ni pool
            // compartido). La sección = frontera inmediatamente ARRIBA (Mes→Oportuno,
            // Aplicacion→Aplicación, sin frontera→1.ª) y el match consume el rótulo de empresa
            // normalizado por el catálogo de sinónimos (SinonimosEmpresaR1, Plan 34). Una fila cuya
            // (firma, sección) no tiene slot en el contrato T0 NO se coloca por posición: el run
            // FALLA con ERR-SIN-SLOT nombrando la fila, su rótulo de empresa y la sección esperada.
            var grupo = grupos.Find(g => EncajaEnSlot(fila, seccionesFuente[i], g.Representante));
            var slot = grupo?.SiguienteLibre();

            if (slot is null)
            {
                throw new CalculoInvalidoException(
                    CodigoError.SinSlot,
                    MensajeSinSlot(bloque.Ase.Id, fila, seccionesFuente[i], ocurrencia, grupo?.UltimoConsumidoF));
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
    /// Slot de dato del contrato: fila del template con su firma A–E y su SECCIÓN (frontera
    /// inmediatamente ARRIBA; solo vinculante para filas-dato de empresa). Las celdas F/L son
    /// referencias A1.
    /// </summary>
    internal sealed record SlotR1(int Fila, string CeldaF, string CeldaL, string A, string B, string C, string D, string E, SeccionFronteraArriba? Seccion);

    /// <summary>
    /// Grupo de slots de la MISMA firma y SECCIÓN (Plan 41/T1, D-A) con su propio cursor ordenado.
    /// Cada (firma, sección) avanza su cursor sin interferir con los demás: un slot-resumen
    /// intercalado antes en el template sigue disponible para una fila que la fuente emite después,
    /// y —clave del cierre E1-2— cada fila cae en la sección que le corresponde por su frontera
    /// inmediatamente ARRIBA (no en el primer slot físico libre). Para filas NO-dato de empresa
    /// (Mes, AFaseo, Componente, Subs/Cont, Total) la sección es <c>null</c>: el grupo es por firma
    /// sola y conserva el comportamiento físico histórico. Plan 41 (Enmienda E-1, E1-1): NO existe
    /// ningún cursor de respaldo — este es el ÚNICO mecanismo de colocación.
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
    }

    /// <summary>
    /// Verdad si la fila fuente y el slot pertenecen a la misma firma + sección (localizador puro):
    /// AFaseo dual-firma primero (cualquiera de las dos variantes), luego A/B/D/E exactos
    /// (OrdinalIgnoreCase) + C por sinonimia; y —solo en filas-dato de empresa— la SECCIÓN
    /// (frontera inmediatamente arriba) debe coincidir. Cero ramas por nombre de empresa (catálogo
    /// abierto por construcción). Plan 41 (T1, D-A): sustituye la clasificación por frontera-debajo
    /// (Plan 35-T2, intacta en <see cref="R1FirmaInterior.EsAplicacionPorFrontera"/>) por la sección.
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

        // Plan 41-T1 (D-A) como parte de la llave en filas-dato de empresa: Oportuno vs Aplicación y
        // 1.ª vs 2.ª ocurrencia NO se mezclan. Si algún lado no clasifica (null: fila no-dato de
        // empresa), la sección no veta (nunca veto ciego para Mes/AFaseo/Componente/SubsCont/Total).
        if (R1FirmaInterior.EsDatoEmpresa(fila)
            && seccionFuente.HasValue
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
    /// AFaseo dual + misma sección en filas-dato de empresa); si no encaja en ninguno, abre un grupo
    /// nuevo. El cursor por grupo de sección es la resolución ÚNICA: cierra los 7 pares E1-2 sin
    /// ramas por empresa ni por período y sin ningún fallback (Enmienda E-1, E1-1).
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
    /// Mensaje de <c>ERR-SIN-SLOT</c> (Plan 41, Enmienda E-1, E1-3): nombra la FILA (firma A|B|C|D|E +
    /// ocurrencia), su RÓTULO DE EMPRESA (columna C) y la SECCIÓN ESPERADA (la resuelta por su
    /// frontera inmediatamente arriba) — información accionable, nunca un fallo mudo ni una colocación
    /// por posición.
    /// </summary>
    private static string MensajeSinSlot(int aseId, FilaEspejoR1 fila, SeccionFronteraArriba? seccion, int ocurrencia, string? celdaCandidata)
    {
        var razon = celdaCandidata is null
            ? "sin ningún slot con firma + sección compatibles en el contrato T0"
            : $"sin slot libre en el contrato T0 (celda candidata: último slot consumido {celdaCandidata})";
        return $"ERR-SIN-SLOT: ASE {aseId}: la fuente R1 (Recaudoporcomponente) trae la fila " +
            $"A|B|C|D|E='{fila.Firma}' con rótulo de empresa '{fila.C}' en la sección esperada " +
            $"'{TipoSeccion(seccion)}' (ocurrencia #{ocurrencia} de esa firma) {razon} del ASE {aseId}. " +
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

    private static IReadOnlyList<SeccionFronteraArriba?> ClasificarSecciones(IReadOnlyList<FilaEspejoR1> filas)
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

        var secciones = ClasificarSecciones(firmas);
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
