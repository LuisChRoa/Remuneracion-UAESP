using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 29 (T3, Unidad R — ESCRITURA): verifica END-TO-END que el desglose-detalle por componente
/// de <c>Rem. Anticipos R2</c> y <c>Reversion Pagos R4</c> se escribe contra la plantilla canónica
/// en ceros del período (julio: <c>Docs/Prueba Julio-2/Plantilla_Remuneracion.xlsx</c>;
/// agosto: base <c>Docs/Prueba Agosto-2/Plantilla_Remuneracion.xlsx</c>, Plan 30/T3), con los
/// insumos REALES de julio-Q2 (<c>Docs/Prueba Julio-2/Insumos</c>) y agosto-Q2
/// (<c>Docs/Prueba Agosto-2/Insumos</c>).
///
/// Cobertura (S3/S4 + R-R-1/R-R-2/R-R-3):
///   - Julio: bloque-destino completo (todas las filas y columnas del mapa) contra el MANUAL del
///     administrativo (misma malla canónica ⇒ oráculo independiente) y la primera fila
///     <c>Vlr Servicio</c> contra la FUENTE leída por el reader.
///   - Agosto: primera fila <c>Vlr Servicio</c> contra la FUENTE (la malla de agosto deriva en
///     LIME/BOGOTA/PROMO/CIUDAD — T0b; por eso no se compara bloque a bloque con el manual, cuya
///     malla de agosto está redimensionada) + la invariante <c>Especiales K</c> del ASE1 (144941.68).
///   - R-R-3: una celda-destino convertida en fórmula en una copia temporal ⇒ ERR-PLANTILLA,
///     sin sobrescritura.
///
/// Sin emojis. Tolerancia ±0.5.
/// </summary>
public sealed class DetalleR2R4EscrituraTests
{
    private const decimal Tol = 0.5m;

    private const string HojaR2 = WorkbookLeafCellMapDetalleR2R4.HojaR2;
    private const string HojaR4 = WorkbookLeafCellMapDetalleR2R4.HojaR4;

    /// <summary>
    /// Plan 59 (H5): base v2 del julio-Q2. Los casos de julio de este fixture citan FILAS de la
    /// geometría v2 (E3/E4/E7/E9, E78..E80, E408..E425 y el bloque 67..124) tomadas de la evidencia
    /// T0b del Plan 29; sobre la v3 esas filas ya no son los mismos slots (los bloques de R2 pasan a
    /// 1/75/196/317/441, §0.4-b) y las cifras citadas seguirían siendo ciertas, pero en otras filas.
    /// La cobertura equivalente sobre la v3 (localización por identidad firma+canal, ambas quincenas,
    /// 5 ASE, contra fuente y oráculo) está en `DetalleR2R4EscritoPlan59Tests`.
    /// </summary>
    private static readonly string PlantillaV2JulioQ2 = Insumos.BaseV2DelPeriodo("2026072");
    private static readonly string ManualJulio = Insumos.ManualJulioQ2;

    private static readonly string ManualAgosto = Insumos.ManualAgosto2026082;

    private static readonly string CarpetaJulio = Insumos.CarpetaInsumosJulioQ2;
    private static readonly string CarpetaAgosto = Insumos.CarpetaInsumosAgosto;

    // ── S3/S4: celdas citadas de la evidencia T0b (julio ASE1) ────────────────────────────────

    [Fact]
    public void JulioAse1_CeldasCitadas_IgualanFuenteYManual()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, PlantillaV2JulioQ2, salida.Ruta);

        // Plan 47 (H1.1, caso testigo autorizado — malla OFICIAL con grupo A/EAB + E/ENEL):
        // el `Vlr Servicio`/`Vlr Intereses` de la fuente (canal E, cerrado por `R7 C=E` de la fuente y
        // por `R9 C=E D=Total` de la plantilla) van a las filas 7/8; el grupo A/EAB (sin fuente) queda
        // en el 0 de la plantilla (E3 = E4 = 0). Antes de H1 esos valores caían en E3/E4 por orden.
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E3") - 0m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E4") - 0m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E7") - 104754634.94m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "F7") - 6975298.69m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "K7") - 0m, -Tol, Tol); // julio sin columna Especiales
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E9") - 104770451.71m, -Tol, Tol);

        // El manual del administrativo usa la malla LEGACY (sin grupo A/EAB): la MISMA cifra vive en
        // SU fila 3. Plan 36/38 ya declaró la obsolescencia de ese oráculo celda-a-celda — se cita, no
        // se compara por celda.
        Assert.InRange(LeerCeldaNumerica(ManualJulio, HojaR2, "E3") - 104754634.94m, -Tol, Tol);

        // R4 ASE1 (sin cambio de slot: la fuente trae los grupos ENEL y OCCIDENTE en el mismo orden que
        // la plantilla): D3 = -15794348.21 (Vlr Servicio de ENEL) y D9 = -762167.27 (Vlr Servicio de OCCIDENTE).
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR4, "D3") - (-15794348.21m), -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR4, "D9") - (-762167.27m), -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(ManualJulio, HojaR4, "D3") - (-15794348.21m), -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(ManualJulio, HojaR4, "D9") - (-762167.27m), -Tol, Tol);
    }

    // ── R-R-1 (Plan 38, T3, R-B-4 / D-C): RETIRADOS con motivo declarado ──────────────────────
    //
    // Los dos tests `JulioQ2_DetalleR2_BloqueCompleto_IgualaManual` / `..._DetalleR4_...`
    // comparaban el bloque-destino completo contra el MANUAL del administrativo de julio-Q2
    // (`Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx`) bajo la
    // premisa "misma malla canónica ⇒ oráculo independiente". El Plan 36/38 cambiaron esa premisa:
    //   (a) la malla del manual es la plantilla LEGACY (otra geometría: p. ej. R2 ASE1 con el grupo
    //       "E"/ENEL en las filas 5-6 y "Componente" en 17, vs el contrato nuevo A/EAB 5-6 y
    //       Componente 21); la comparación celda-a-celda es, por tanto, contra un oráculo obsoleto;
    //   (b) además, el manual de R4 ASE2 trae un bloque `NUEVO ESQUEMA` (D104 = -290750) que la
    //       fuente Q2 ACTUAL ya no contiene — diferencia de FUENTE, no de escritura.
    //
    // La cobertura equivalente (bloque-destino == FUENTE leída por el reader, celda a celda) vive en
    // `*_PrimeraVlrServicio_IgualaFuente` (los 5 ASE × R2/R4 × julio/agosto) y en
    // `JulioAse1_CeldasCitadas_IgualanFuenteYManual`. NO se relaja ningún assert: se retira el oráculo
    // inválido con motivo citado (D-C); la regresión contra el manual queda como dependencia declarada
    // de T5 (comparador con brechas versionadas).

    // ── R-R-1/R-R-2: julio y agosto, primera fila Vlr Servicio vs fuente (reader) ─────────────

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void JulioQ2_DetalleR2_PrimeraVlrServicio_IgualaFuente(int aseId)
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, PlantillaV2JulioQ2, salida.Ruta);

        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var fuente = reader.LeerDetalleR2(Insumos.Ase(aseId), Insumos.R2JulioQ2(aseId));
        AssertFilaVlrServicioR2(salida.Ruta, aseId, fuente, $"julio ASE{aseId}");
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void JulioQ2_DetalleR4_PrimeraVlrServicio_IgualaFuente(int aseId)
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, PlantillaV2JulioQ2, salida.Ruta);

        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var fuente = reader.LeerDetalleR4(Insumos.Ase(aseId), Insumos.R4JulioQ2(aseId));
        AssertFilaVlrServicioR4(salida.Ruta, aseId, fuente, $"julio ASE{aseId}");
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void AgostoQ2_DetalleR2_PrimeraVlrServicio_IgualaFuente(int aseId)
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaAgosto, new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 }, Insumos.PlantillaAgosto2026082, salida.Ruta);

        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var fuente = reader.LeerDetalleR2(Insumos.Ase(aseId), Insumos.R2AgostoQ2(aseId));
        AssertFilaVlrServicioR2(salida.Ruta, aseId, fuente, $"agosto ASE{aseId}");
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void AgostoQ2_DetalleR4_PrimeraVlrServicio_IgualaFuente(int aseId)
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaAgosto, new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 }, Insumos.PlantillaAgosto2026082, salida.Ruta);

        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var fuente = reader.LeerDetalleR4(Insumos.Ase(aseId), Insumos.R4AgostoQ2(aseId));
        AssertFilaVlrServicioR4(salida.Ruta, aseId, fuente, $"agosto ASE{aseId}");
    }

    /// <summary>
    /// Agosto ASE1: la fuente trae <c>Especiales</c> en K y el destino la escribe (invariante
    /// de la columna ausente en julio). Plan 47 (H1.1): el `Vlr Servicio` (canal E, cierre
    /// `R9 C=E D=Total`) va a la FILA 7 — no a la 3, que pertenece al grupo A/EAB sin fuente en la
    /// fuente — con E7 = 68638686.24 / F7 = 3508770.06 / K7 = 144941.68 y E3 = 0.
    /// </summary>
    [Fact]
    public void AgostoAse1_EspecialesEnK_IgualaFuenteYManual()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaAgosto, new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 }, Insumos.PlantillaAgosto2026082, salida.Ruta);

        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E3") - 0m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E7") - 68638686.24m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "F7") - 3508770.06m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "K7") - 144941.68m, -Tol, Tol);

        // Manual de agosto (malla LEGACY, sin grupo A/EAB): la misma cifra vive en SU fila 3.
        Assert.True(File.Exists(ManualAgosto), $"Falta el manual de agosto: {ManualAgosto}");
        Assert.InRange(LeerCeldaNumerica(ManualAgosto, HojaR2, "E3") - 68638686.24m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(ManualAgosto, HojaR2, "K3") - 144941.68m, -Tol, Tol);
    }

    // ── R-R-3: destino-detalle fórmula ⇒ ERR-PLANTILLA (copia temp, nunca sobrescritura) ──────

    [Fact]
    public void DetalleR2_DestinoFormula_LanzaErrPlantilla_SinSobrescribir()
    {
        using var temp = new DirectorioTemporal("remuneracion-detalle-formula-");
        var plantillaMutada = Path.Combine(temp.Ruta, "plantilla-mutada.xlsx");
        File.Copy(PlantillaV2JulioQ2, plantillaMutada);

        // E7 = primera celda de captura del detalle R2 del ASE1 que el writer SÍ escribe (Plan 47 H1.1:
        // el `Vlr Servicio` de la fuente, canal E cerrado por `R9 C=E D=Total`, cae en la fila 7): se
        // convierte en fórmula.
        ConvertirEnFormula(plantillaMutada, HojaR2, "E7", "1+1");

        var salida = Path.Combine(temp.Ruta, "salida.xlsx");
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, plantillaMutada, salida));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("fórmula", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ── Plan 47 W-2 / W-3: fail-fast del detalle R2/R4 y robustez al orden ─────────────────────

    /// <summary>
    /// W-2 (género de <c>EscritorValoresSlotR1Tests.S3_EmpresaSinSlot_FailFastSinRemap</c>): una fila
    /// SINTÉTICA de la fuente cuya clave <c>(firma + canal efectivo)</c> NO está en la lista de ceros
    /// documentados hace fail-fast <c>ERR-SIN-SLOT</c> con mensaje accionable (hoja + ASE + nº de fila +
    /// firma + canal), sin remap ni colocación por posición, y sin dejar salida parcial.
    /// </summary>
    [Fact]
    public void DetalleR2_FilaSinteticaSinSlot_LanzaErrSinSlot_ConMensajeAccionable()
    {
        using var temp = new DirectorioTemporal("remuneracion-detalle-sinslot-");
        var insumos = CopiarInsumos(temp.Ruta);
        var fuente = Directory.EnumerateFiles(Path.Combine(insumos, "2-Lime"), "*.xlsx")
            .First(f => Path.GetFileName(f).StartsWith("RerpoteDetalleSaldosaFavor", StringComparison.OrdinalIgnoreCase));
        InsertarGrupoSintetico(fuente, concepto: "Vlr Servicio", canal: "Z");

        var salida = Path.Combine(temp.Ruta, "salida.xlsx");
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            EjecutarFlujo(insumos, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, PlantillaV2JulioQ2, salida));

        Assert.Equal(CodigoError.SinSlot, ex.Codigo);
        Assert.Contains("ERR-SIN-SLOT", ex.Message, StringComparison.Ordinal);
        Assert.Contains("'Rem. Anticipos R2'", ex.Message, StringComparison.Ordinal);
        Assert.Contains("ASE 2", ex.Message, StringComparison.Ordinal);
        Assert.Contains("fila fuente #", ex.Message, StringComparison.Ordinal);
        Assert.Contains("firma '|||Vlr Servicio'", ex.Message, StringComparison.Ordinal);
        Assert.Contains("canal efectivo 'Z'", ex.Message, StringComparison.Ordinal);
        Assert.Contains("bloque 67..124", ex.Message, StringComparison.Ordinal);
        Assert.Contains("ceros documentados", ex.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(salida), "el parcial debe borrarse (atomicidad).");
    }

    /// <summary>
    /// Plan 48 (H1.1 — T0 §2.2): el `Vlr Intereses` del canal 2 de ASE2 era un "cero documentado" del
    /// Plan 47; la geometría v2 SÍ tiene slot (R79 dentro del bloque 67..124) y el valor se escribe.
    /// El falso hueco quedó cerrado: no se omite la fila ni se pierde el cierre `||2|Total` (R80).
    /// </summary>
    [Fact]
    public void DetalleR2_FalsoHuecoCerrado_ViCanal2_SeEscribeEnSuSlot()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, PlantillaV2JulioQ2, salida.Ruta);

        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E78") - 1616173.63m, -Tol, Tol); // `Vlr Servicio` canal 2/NS
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E79") - 406.42m, -Tol, Tol);     // `Vlr Intereses` canal 2/NS (antes omitido)
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E80") - 1616580.05m, -Tol, Tol); // cierre `||2|Total`
    }

    /// <summary>
    /// W-3 (prueba de que el último camino posicional murió): la fuente de ASE5 con los DOS grupos del
    /// canal <c>E</c> (ENEL y ENERBIT) en orden INVERTIDO. Con el canal compuesto por segundo cierre
    /// (<c>E/ENEL</c> vs <c>E/ENERBIT</c>) cada grupo recibe SU dato en SU slot (386..389 / 390..393),
    /// independientemente del orden de la fuente.
    /// </summary>
    [Fact]
    public void DetalleR2_Ase5_DobleGrupoE_OrdenInvertido_CadaGrupoRecibeLoSuyo()
    {
        using var temp = new DirectorioTemporal("remuneracion-detalle-orden-");
        var insumos = CopiarInsumos(temp.Ruta);
        var fuente = Directory.EnumerateFiles(Path.Combine(insumos, "5-Área Limpia"), "*.xlsx")
            .First(f => Path.GetFileName(f).StartsWith("RerpoteDetalleSaldosaFavor", StringComparison.OrdinalIgnoreCase));
        InvertirContenidoGruposCanalE(fuente);

        var salida = Path.Combine(temp.Ruta, "salida-orden.xlsx");
        EjecutarFlujo(insumos, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, PlantillaV2JulioQ2, salida);

        // Plan 48 (H1.1 — T0 §1.3): la geometría v2 desplaza ASE5 +22; los grupos `E/ENEL` (408..410)
        // y `E/ENERBIT` (412..414) conservan el reparto por canal compuesto aunque la fuente los traiga
        // al revés; OCCIDENTE (425) no se contagia.
        Assert.InRange(LeerCeldaNumerica(salida, HojaR2, "E408") - 19713558.23m, -Tol, Tol); // ENEL Vlr Servicio
        Assert.InRange(LeerCeldaNumerica(salida, HojaR2, "E410") - 19732312.24m, -Tol, Tol); // ENEL cierre `||E|Total`
        Assert.InRange(LeerCeldaNumerica(salida, HojaR2, "E412") - 7522088.94m, -Tol, Tol);  // ENERBIT Vlr Servicio
        Assert.InRange(LeerCeldaNumerica(salida, HojaR2, "E414") - 7529567.64m, -Tol, Tol);  // ENERBIT cierre `||E|Total`
        Assert.InRange(LeerCeldaNumerica(salida, HojaR2, "E425") - 11900986.83m, -Tol, Tol); // OCCIDENTE (sin contagio)
    }

    /// <summary>Copia el árbol de insumos de julio-2 a un temp (nunca se muta el repo).</summary>
    private static string CopiarInsumos(string destino)
    {
        var origen = Insumos.CarpetaInsumosJulioQ2;
        var copia = Path.Combine(destino, "Insumos");
        foreach (var dir in Directory.GetDirectories(origen, "*", SearchOption.AllDirectories))
        {
            Directory.CreateDirectory(dir.Replace(origen, copia, StringComparison.OrdinalIgnoreCase));
        }

        Directory.CreateDirectory(copia);
        foreach (var archivo in Directory.GetFiles(origen, "*", SearchOption.AllDirectories))
        {
            File.Copy(archivo, archivo.Replace(origen, copia, StringComparison.OrdinalIgnoreCase), overwrite: true);
        }

        return copia;
    }

    /// <summary>Inserta, antes del cierre <c>A=Total</c> del bloque, un grupo sintético
    /// (<c>D=&lt;concepto&gt;</c> + cierre <c>C=&lt;canal&gt; D=Total</c>): clave fuera de la lista de
    /// ceros documentados.</summary>
    private static void InsertarGrupoSintetico(string ruta, string concepto, string canal)
    {
        using var doc = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, "Sheet1");
        var sheetData = worksheet.GetFirstChild<SheetData>() ?? throw new InvalidOperationException("SheetData null");
        var cierre = worksheet.Descendants<Row>()
            .Last(r => string.Equals(TextoCelda(workbookPart, worksheet, "A" + r.RowIndex!.Value), "Total", StringComparison.OrdinalIgnoreCase));
        var numero = cierre.RowIndex!.Value;

        foreach (var fila in worksheet.Descendants<Row>().Where(r => r.RowIndex!.Value >= numero).ToList())
        {
            fila.RowIndex = fila.RowIndex!.Value + 2;
        }

        var filaConcepto = new Row { RowIndex = numero };
        filaConcepto.Append(CeldaTexto("D", numero, concepto), CeldaNumero("E", numero, 1m));
        var filaCierre = new Row { RowIndex = numero + 1 };
        filaCierre.Append(CeldaTexto("C", numero + 1, canal), CeldaTexto("D", numero + 1, "Total"));

        sheetData.InsertBefore(filaConcepto, cierre);
        sheetData.InsertBefore(filaCierre, cierre);
        worksheet.Save();
    }

    /// <summary>Intercambia el CONTENIDO (no la posición) de los dos grupos del canal <c>E</c> de la
    /// fuente R2 de ASE5 — grupos de 4 filas: concepto, concepto, cierre `C=E`, empresa.</summary>
    private static void InvertirContenidoGruposCanalE(string ruta)
    {
        using var doc = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, "Sheet1");
        var cierres = worksheet.Descendants<Row>()
            .Where(r => string.Equals(TextoCelda(workbookPart, worksheet, "C" + r.RowIndex!.Value), "E", StringComparison.OrdinalIgnoreCase)
                     && string.Equals(TextoCelda(workbookPart, worksheet, "D" + r.RowIndex!.Value), "Total", StringComparison.OrdinalIgnoreCase))
            .OrderBy(r => r.RowIndex!.Value)
            .Take(2) // los dos grupos E de la MITAD Anticipos (la mitad Subs/Cont tiene otros dos)
            .ToList();
        Assert.Equal(2, cierres.Count);

        var inicio1 = cierres[0].RowIndex!.Value - 2;
        var inicio2 = cierres[1].RowIndex!.Value - 2;
        foreach (var columna in new[] { "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P" })
        {
            for (uint k = 0; k < 4; k++)
            {
                IntercambiarContenido(workbookPart, worksheet, columna, inicio1 + k, columna, inicio2 + k);
            }
        }

        worksheet.Save();
    }

    private static void IntercambiarContenido(WorkbookPart workbookPart, Worksheet worksheet, string columna, uint filaA, string columnaB, uint filaB)
    {
        var a = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, columna + filaA.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase));
        var b = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, columnaB + filaB.ToString(CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase));
        if (a is null || b is null)
        {
            return;
        }

        var contenidoA = Contenido(a);
        var contenidoB = Contenido(b);
        Aplicar(a, contenidoB);
        Aplicar(b, contenidoA);
    }

    private static (CellValues? Tipo, List<OpenXmlElement> Hijos) Contenido(Cell celda) =>
        (celda.DataType?.Value, celda.ChildElements.Select(e => (OpenXmlElement)e.CloneNode(true)).ToList());

    private static void Aplicar(Cell celda, (CellValues? Tipo, List<OpenXmlElement> Hijos) contenido)
    {
        celda.RemoveAllChildren();
        celda.DataType = contenido.Tipo;
        foreach (var hijo in contenido.Hijos)
        {
            celda.AppendChild(hijo);
        }
    }

    private static Cell CeldaTexto(string columna, uint fila, string texto) => new()
    {
        CellReference = columna + fila.ToString(CultureInfo.InvariantCulture),
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(texto)),
    };

    private static Cell CeldaNumero(string columna, uint fila, decimal valor) => new()
    {
        CellReference = columna + fila.ToString(CultureInfo.InvariantCulture),
        CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture)),
    };

    // ── Helpers ──────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Plan 47 (H1.1/H1.2): la primera fila <c>Vlr Servicio</c> de la FUENTE debe caer en el slot
    /// cuya clave <c>(firma + canal efectivo)</c> coincide — el canal efectivo se calcula con el
    /// predicado PURO de Core (<see cref="CanalEfectivoDetalleR2R4"/>, R46-1 análogo) sobre los
    /// rótulos del bloque ya escrito y sobre las filas de la fuente leídas por el reader.
    /// Se verifica además que el valor NO se cuele en la primera fila del bloque cuando su canal
    /// efectivo difiere (casamiento por orden eliminado, Plan 41 E1-1).
    /// </summary>
    private static void AssertFilaVlrServicioR2(string salida, int aseId, DetalleR2AseInputs fuente, string contexto)
    {
        var fila = fuente.Filas.FirstOrDefault(f => f.Firma == "|||Vlr Servicio")
            ?? throw new InvalidOperationException($"{contexto}: la fuente no trae la fila 'Vlr Servicio'.");
        var rotulosFuente = fuente.Filas.Select(f => new CanalEfectivoDetalleR2R4.Rotulo(f.A, f.B, f.C, f.D)).ToList();
        var canal = CanalEfectivoDetalleR2R4.CanalEfectivo(
            rotulosFuente, 0, WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR2, WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR2);

        var bloque = BloqueDelDetalle(salida, HojaR2, DelimitadorBloquesHoja.R2, aseId);
        var destino = LocalizarSlot(
            salida, HojaR2, bloque, "|||Vlr Servicio", canal,
            WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR2,
            WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR2, [0, 1, 2, 3]);

        // Plan 61 (H1.6, D9): el destino de la columna se resuelve contra la fila de encabezados que la
        // plantilla declara en ESTA hoja, no contra una lista congelada. Se verifica cada componente de
        // la fuente y, además, que ninguna columna declarada con monto quede sin escribir.
        AssertComponentesEscritos(salida, HojaR2, DelimitadorBloquesHoja.R2, aseId, destino, fila.ValoresPorComponente, contexto);

        var primerCanal = CanalEfectivoDetalleR2R4.CanalEfectivo(
            RotulosDestino(salida, HojaR2, bloque), 0,
            WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR2, WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR2);
        if (!string.Equals(primerCanal, canal, StringComparison.OrdinalIgnoreCase))
        {
            Assert.InRange(
                LeerCeldaNumerica(salida, HojaR2, "E" + bloque.FilaInicio.ToString(CultureInfo.InvariantCulture)), -Tol, Tol);
        }
    }

    private static void AssertFilaVlrServicioR4(string salida, int aseId, DetalleR4AseInputs fuente, string contexto)
    {
        var fila = fuente.Filas.FirstOrDefault(f => f.C == "Vlr Servicio")
            ?? throw new InvalidOperationException($"{contexto}: la fuente no trae la fila 'Vlr Servicio'.");
        var rotulosFuente = fuente.Filas
            .Select(f => new CanalEfectivoDetalleR2R4.Rotulo(f.A, f.B, f.C, string.Empty))
            .ToList();
        var canal = CanalEfectivoDetalleR2R4.CanalEfectivo(
            rotulosFuente, 0, WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR4, WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR4);

        var bloque = BloqueDelDetalle(salida, HojaR4, DelimitadorBloquesHoja.R4, aseId);
        var destino = LocalizarSlot(
            salida, HojaR4, bloque, "||Vlr Servicio", canal,
            WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR4,
            WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR4, [0, 1, 2]);

        AssertComponentesEscritos(salida, HojaR4, DelimitadorBloquesHoja.R4, aseId, destino, fila.ValoresPorComponente, contexto);

        var primerCanal = CanalEfectivoDetalleR2R4.CanalEfectivo(
            RotulosDestino(salida, HojaR4, bloque), 0,
            WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR4, WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR4);
        if (!string.Equals(primerCanal, canal, StringComparison.OrdinalIgnoreCase))
        {
            Assert.InRange(
                LeerCeldaNumerica(salida, HojaR4, "D" + bloque.FilaInicio.ToString(CultureInfo.InvariantCulture)), -Tol, Tol);
        }
    }

    /// <summary>
    /// Plan 61 (H1.6, D9): verifica la fila escrita componente a componente resolviendo la COLUMNA
    /// contra la fila de encabezados que la propia hoja declara (la autoridad del destino). Un
    /// componente de la fuente con monto que no resuelva a una columna declarada es una pérdida y el
    /// assert la nombra; una columna declarada que no recibió su valor también.
    /// </summary>
    private static void AssertComponentesEscritos(
        string salida,
        string hoja,
        DelimitadorBloquesHoja.HojaDelimitable hojaDelimitable,
        int aseId,
        string filaDestino,
        IReadOnlyDictionary<string, decimal?> valoresFuente,
        string contexto)
    {
        using var workbook = SpreadsheetDocument.Open(salida, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, hoja);
        var identidad = $"hoja '{hoja}' · ASE {aseId}";
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hojaDelimitable, aseId, worksheet);
        var columnas = ResolvedorDestinoDetalleR2R4.ColumnasDeclaradas(workbookPart, worksheet, bloque.EncabezadoColumnas, identidad);
        var indice = ResolvedorDestinoDetalleR2R4.IndicePorEtiqueta(columnas, identidad);

        var esperadoPorColumna = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        foreach (var (encabezado, valor) in valoresFuente)
        {
            var columna = ResolvedorDestinoDetalleR2R4.ColumnaDe(indice, encabezado);
            Assert.True(
                columna is not null || valor is null || valor.Value == 0m,
                $"{contexto}: el componente '{encabezado}' trae {valor} y la plantilla no declara columna para él ({identidad}).");
            if (columna is not null)
            {
                esperadoPorColumna[columna] = valor ?? 0m;
            }
        }

        Assert.NotEmpty(columnas);
        foreach (var (etiqueta, columna) in columnas)
        {
            var esperado = esperadoPorColumna.GetValueOrDefault(columna);
            var real = LeerCeldaNumerica(salida, hoja, columna + filaDestino);
            Assert.True(
                Math.Abs(real - esperado) <= Tol,
                $"{contexto}: {hoja} {columna}{filaDestino} ('{etiqueta}') escrito={real} fuente={esperado} (Δ={real - esperado}).");
        }
    }

    /// <summary>
    /// Bloque de detalle del ASE en la hoja, DELIMITADO por su cabecera y su gran total declarados
    /// (Plan 59 §11-bis). Antes se tomaba del mapa congelado de la v2 (<c>BloquesR2/R4</c>): con la
    /// plantilla v3 ese rango ya no es el del bloque y el slot se buscaba fuera de él.
    /// </summary>
    private static (int FilaInicio, int FilaFin) BloqueDelDetalle(
        string salida,
        string hoja,
        DelimitadorBloquesHoja.HojaDelimitable hojaDelimitable,
        int aseId)
    {
        using var workbook = SpreadsheetDocument.Open(salida, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hojaDelimitable, aseId);
        return (bloque.EncabezadoColumnas + 1, bloque.GranTotal);
    }

    /// <summary>Rótulos A–D del bloque tal como quedaron en el workbook generado (resolviendo las
    /// cadenas compartidas: `Cell.InnerText` devolvería el ÍNDICE de la shared string, no el texto).</summary>
    private static IReadOnlyList<CanalEfectivoDetalleR2R4.Rotulo> RotulosDestino(string salida, string hoja, (int FilaInicio, int FilaFin) bloque)
    {
        using var workbook = SpreadsheetDocument.Open(salida, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, hoja);
        var rotulos = new List<CanalEfectivoDetalleR2R4.Rotulo>(bloque.FilaFin - bloque.FilaInicio + 1);
        for (var fila = bloque.FilaInicio; fila <= bloque.FilaFin; fila++)
        {
            rotulos.Add(new CanalEfectivoDetalleR2R4.Rotulo(
                TextoCelda(workbookPart, worksheet, "A" + fila.ToString(CultureInfo.InvariantCulture)),
                TextoCelda(workbookPart, worksheet, "B" + fila.ToString(CultureInfo.InvariantCulture)),
                TextoCelda(workbookPart, worksheet, "C" + fila.ToString(CultureInfo.InvariantCulture)),
                TextoCelda(workbookPart, worksheet, "D" + fila.ToString(CultureInfo.InvariantCulture))));
        }

        return rotulos;
    }

    private static string TextoCelda(WorkbookPart workbookPart, Worksheet worksheet, string celda)
    {
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        if (objCelda is null)
        {
            return string.Empty;
        }

        if (objCelda.DataType?.Value == CellValues.SharedString
            && int.TryParse(objCelda.CellValue?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice))
        {
            var tabla = workbookPart.SharedStringTablePart?.SharedStringTable;
            var item = tabla?.Elements<SharedStringItem>().ElementAtOrDefault(indice);
            return item?.InnerText?.Trim() ?? string.Empty;
        }

        return objCelda.CellValue?.InnerText?.Trim() ?? string.Empty;
    }

    /// <summary>Fila (número) del slot cuya clave (firma + canal efectivo) coincide con la pedida.</summary>
    private static string LocalizarSlot(
        string salida,
        string hoja,
        (int FilaInicio, int FilaFin) bloque,
        string firma,
        string? canal,
        IReadOnlyList<int> columnasEtiquetaPropia,
        IReadOnlyList<int> columnasGrupo,
        IReadOnlyList<int> columnasFirma)
    {
        var rotulos = RotulosDestino(salida, hoja, bloque);
        for (var i = 0; i < rotulos.Count; i++)
        {
            var firmaDestino = string.Join("|", columnasFirma.Select(c => rotulos[i].Texto(c).Trim()));
            var canalDestino = CanalEfectivoDetalleR2R4.CanalEfectivo(rotulos, i, columnasEtiquetaPropia, columnasGrupo);
            if (string.Equals(firmaDestino, firma, StringComparison.OrdinalIgnoreCase)
                && string.Equals(canalDestino, canal, StringComparison.OrdinalIgnoreCase))
            {
                return (bloque.FilaInicio + i).ToString(CultureInfo.InvariantCulture);
            }
        }

        throw new InvalidOperationException(
            $"No se encontró slot para la firma '{firma}' con canal efectivo '{canal ?? "(sin canal)"}' en {hoja} {bloque.FilaInicio}..{bloque.FilaFin}.");
    }

    private static void EjecutarFlujo(string carpetaInsumos, Periodo periodo, string rutaPlantilla, string rutaSalida)
    {
        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());

        procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = periodo,
            CarpetaPeriodo = carpetaInsumos,
            RutaPlantilla = rutaPlantilla,
            RutaSalida = rutaSalida
        });
    }

    private static decimal LeerCeldaNumerica(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        if (objCelda.CellValue is null || string.IsNullOrWhiteSpace(objCelda.CellValue.InnerText))
        {
            return 0m;
        }

        return decimal.TryParse(objCelda.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor)
            ? valor
            : 0m;
    }

    private static void ConvertirEnFormula(string ruta, string hoja, string celda, string formula)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var worksheet = Hoja(documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null"), hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        objCelda.CellValue?.Remove();
        objCelda.DataType = null;
        objCelda.CellFormula = new CellFormula(formula);
        worksheet.Save();
    }

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet.");
    }

    private static SalidaTemporal NuevaSalida() => new("remuneracion-detalle-" + Guid.NewGuid().ToString("N"));

    private sealed class SalidaTemporal : IDisposable
    {
        private readonly DirectorioTemporal _dir;

        public SalidaTemporal(string prefijo)
        {
            _dir = new DirectorioTemporal(prefijo);
            Ruta = Path.Combine(_dir.Ruta, "salida.xlsx");
        }

        public string Ruta { get; }

        public void Dispose() => _dir.Dispose();
    }

    private sealed class DirectorioTemporal : IDisposable
    {
        public DirectorioTemporal(string prefijo)
        {
            Ruta = Path.Combine(Path.GetTempPath(), prefijo + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Ruta);
        }

        public string Ruta { get; }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(Ruta))
                {
                    Directory.Delete(Ruta, recursive: true);
                }
            }
            catch
            {
                // best-effort
            }
        }
    }
}
