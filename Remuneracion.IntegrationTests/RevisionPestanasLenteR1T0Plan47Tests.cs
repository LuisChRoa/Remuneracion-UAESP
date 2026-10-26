using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;
using Rotulo = Remuneracion.Core.Models.CanalEfectivoDetalleR2R4.Rotulo;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 47 — T0/contrato de las "demás pestañas de pegado" bajo la lente del frente R1
/// (fuente / canal / concepto / componentes). Read-only: BCL <see cref="ZipArchive"/> + <see cref="XDocument"/>,
/// sin Excel/COM ni OpenXml SDK en la derivación.
///
/// <para>
/// <b>H0 (2026-10-17)</b> congeló el ESTADO OBSERVADO (mapeo por orden vía LCS) y probó los desvíos:
/// corrimiento en R2 ASE1/2/4/5 (el grupo E caía en las filas del grupo A/EAB), corrimiento en R4 ASE2
/// (OCCIDENTE bajo NUEVO ESQUEMA), falsos huecos (slots VS/VI en 0 con contraparte real) y la omisión
/// silenciosa de filas sin slot. Evidencia: <c>Docs/T0-Plan47-Revision-Pestanas-LenteR1.md</c> §4/§6.
/// </para>
///
/// <para>
/// <b>H1 (autorizado)</b> reemplazó la clave de colocación por <c>(firma + CANAL EFECTIVO)</c>
/// (R46-1 análogo, <see cref="CanalEfectivoDetalleR2R4"/>, puro en Core) + lista de ceros documentados
/// (<see cref="CerosDocumentadosDetalleR2R4"/>) + <c>Otros → R</c> en R4. Este fixture se ACTUALIZÓ con
/// evidencia citada (R-FIXTURE): cada par congelado cita la fila-fuente, el slot y el cierre que lo
/// ubica; los tests de P2 pasaron de "congelar el corrimiento" a "verificar la identidad de canal".
/// </para>
///
/// El OFICIAL nunca se modifica (SHA256 antes/después); la medición contra las corridas 9 y nueva vive
/// en el documento T0 y en el reporte de H2.
/// </summary>
public sealed class RevisionPestanasLenteR1T0Plan47Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string Sha256Oficial = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";

    private const string HojaR2 = "Rem. Anticipos R2";
    private const string HojaR4 = "Reversion Pagos R4";

    private static readonly string[] ColsR2 = ["A", "B", "C", "D"];
    private static readonly string[] ColsR4 = ["A", "B", "C"];

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    /// <summary>
    /// Bloque congelado: firma de cada fila-destino, firma de cada fila-fuente, el mapeo
    /// (fila-destino absoluta ← índice-fuente) que produce la clave <c>(firma + canal efectivo)</c>,
    /// las filas-destino sin fuente y las filas-fuente sin slot (ceros documentados: índice, firma,
    /// canal efectivo, Total de la columna <c>Total</c>).
    /// </summary>
    private sealed record Bloque(
        string Hoja,
        int Inicio,
        int Fin,
        string DestinoFirmas,
        string FuenteFirmas,
        string Pares,
        int[] SinFuente,
        (int Indice, string Firma, string? Canal, decimal Total)[] CerosDocumentados);

    // ── P6 (H1): matriz `Rem. Anticipos R2` — bloques de WorkbookLeafCellMapDetalleR2R4.BloquesR2 ──

    private static readonly Bloque[] MatrizR2 =
    [
        // ASE1 (PROMOAMBIENTAL): la fuente no trae el grupo A/EAB → 3..6 quedan en el 0 de la
        // plantilla; el par del canal E (`R7..R8`, cerrado por `R9 C=E D=Total`) recibe src#0/#1
        // (`Vlr Servicio` 104.754.634,94 / `Vlr Intereses` 15.816,77 — H0 §4 testigo de H1).
        new(HojaR2, 3, 33,
            "|||Vlr Servicio;|||Vlr Intereses;||A|Total;|EAB|Total|;|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENEL|Total|;|||Vlr Servicio;|||Vlr Intereses;||H|Total;|||Vlr Servicio;|||Vlr Intereses;||O|Total;|||Vlr Servicio;|||Vlr Intereses;||T|Total;|OCCIDENTE|Total|;Componente|Total||;|||Subsidio(-)/Contribucion(+);||E|Total;|ENEL|Total|;|||Subsidio(-)/Contribucion(+);||H|Total;|||Subsidio(-)/Contribucion(+);||O|Total;|||Subsidio(-)/Contribucion(+);||T|Total;|OCCIDENTE|Total|;Subs/Cont|Total||;Total|||",
            "|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENEL|Total|;|||Vlr Servicio;|||Vlr Intereses;||H|Total;|||Vlr Servicio;|||Vlr Intereses;||O|Total;|||Vlr Servicio;|||Vlr Intereses;||T|Total;|OCCIDENTE|Total|;Componente|Total||;|||Subsidio(-)/Contribucion(+);||E|Total;|ENEL|Total|;|||Subsidio(-)/Contribucion(+);||H|Total;|||Subsidio(-)/Contribucion(+);||O|Total;|||Subsidio(-)/Contribucion(+);||T|Total;|OCCIDENTE|Total|;Subs/Cont|Total||;Total|||",
            "7<-0,8<-1,9<-2,10<-3,11<-4,12<-5,13<-6,14<-7,15<-8,16<-9,17<-10,18<-11,19<-12,20<-13,21<-14,22<-15,23<-16,24<-17,25<-18,26<-19,27<-20,28<-21,29<-22,30<-23,31<-24,32<-25,33<-26",
            [3, 4, 5, 6], []),

        // ASE2 (LIME): filas 67..70 sin fuente (grupo A/EAB); ceros documentados: `Vlr Intereses` del
        // canal 2 (el grupo trae una sola fila: 78) y el grupo canal 9 (sin fila en el bloque).
        new(HojaR2, 67, 113,
            "|||Vlr Servicio;|||Vlr Intereses;||A|Total;|EAB|Total|;|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENEL|Total|;|||Vlr Servicio;|||Vlr Intereses;||1|Total;|||Vlr Servicio;||2|Total;|||Vlr Servicio;|||Vlr Intereses;||5|Total;|NUEVO ESQUEMA|Total|;|||Vlr Servicio;|||Vlr Intereses;||H|Total;|||Vlr Servicio;|||Vlr Intereses;||O|Total;|||Vlr Servicio;|||Vlr Intereses;||T|Total;|OCCIDENTE|Total|;Componente|Total||;|||Subsidio(-)/Contribucion(+);||E|Total;|ENEL|Total|;|||Subsidio(-)/Contribucion(+);||1|Total;|||Subsidio(-)/Contribucion(+);||2|Total;|||Subsidio(-)/Contribucion(+);||5|Total;|NUEVO ESQUEMA|Total|;|||Subsidio(-)/Contribucion(+);||H|Total;|||Subsidio(-)/Contribucion(+);||O|Total;|||Subsidio(-)/Contribucion(+);||T|Total;|OCCIDENTE|Total|;Subs/Cont|Total||;Total|||",
            "|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENEL|Total|;|||Vlr Servicio;|||Vlr Intereses;||1|Total;|||Vlr Servicio;|||Vlr Intereses;||2|Total;|||Vlr Servicio;|||Vlr Intereses;||5|Total;|||Vlr Servicio;||9|Total;|NUEVO ESQUEMA|Total|;|||Vlr Servicio;|||Vlr Intereses;||H|Total;|||Vlr Servicio;|||Vlr Intereses;||O|Total;|||Vlr Servicio;|||Vlr Intereses;||T|Total;|OCCIDENTE|Total|;Componente|Total||;|||Subsidio(-)/Contribucion(+);||E|Total;|ENEL|Total|;|||Subsidio(-)/Contribucion(+);||1|Total;|||Subsidio(-)/Contribucion(+);||2|Total;|||Subsidio(-)/Contribucion(+);||5|Total;|NUEVO ESQUEMA|Total|;|||Subsidio(-)/Contribucion(+);||H|Total;|||Subsidio(-)/Contribucion(+);||O|Total;|||Subsidio(-)/Contribucion(+);||T|Total;|OCCIDENTE|Total|;Subs/Cont|Total||;Total|||",
            "71<-0,72<-1,73<-2,74<-3,75<-4,76<-5,77<-6,78<-7,79<-9,80<-10,81<-11,82<-12,83<-15,84<-16,85<-17,86<-18,87<-19,88<-20,89<-21,90<-22,91<-23,92<-24,93<-25,94<-26,95<-27,96<-28,97<-29,98<-30,99<-31,100<-32,101<-33,102<-34,103<-35,104<-36,105<-37,106<-38,107<-39,108<-40,109<-41,110<-42,111<-43,112<-44,113<-45",
            [67, 68, 69, 70],
            [
                (8, "|||Vlr Intereses", "2/NUEVO ESQUEMA", 406.42m),
                (13, "|||Vlr Servicio", "9/NUEVO ESQUEMA", 831919.29m),
                (14, "||9|Total", "9/NUEVO ESQUEMA", 831919.29m),
            ]),

        // ASE3 (CIUDAD LIMPIA): la fuente SÍ trae el grupo A/EAB → identidad total (sin ceros).
        new(HojaR2, 164, 194,
            "|||Vlr Servicio;|||Vlr Intereses;||A|Total;|EAB|Total|;|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENEL|Total|;|||Vlr Servicio;|||Vlr Intereses;||H|Total;|||Vlr Servicio;|||Vlr Intereses;||O|Total;|||Vlr Servicio;|||Vlr Intereses;||T|Total;|OCCIDENTE|Total|;Componente|Total||;|||Subsidio(-)/Contribucion(+);||E|Total;|ENEL|Total|;|||Subsidio(-)/Contribucion(+);||H|Total;|||Subsidio(-)/Contribucion(+);||O|Total;|||Subsidio(-)/Contribucion(+);||T|Total;|OCCIDENTE|Total|;Subs/Cont|Total||;Total|||",
            "|||Vlr Servicio;||A|Total;|EAB|Total|;|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENEL|Total|;|||Vlr Servicio;|||Vlr Intereses;||H|Total;|||Vlr Servicio;|||Vlr Intereses;||O|Total;|||Vlr Servicio;|||Vlr Intereses;||T|Total;|OCCIDENTE|Total|;Componente|Total||;|||Subsidio(-)/Contribucion(+);||E|Total;|ENEL|Total|;|||Subsidio(-)/Contribucion(+);||H|Total;|||Subsidio(-)/Contribucion(+);||O|Total;|||Subsidio(-)/Contribucion(+);||T|Total;|OCCIDENTE|Total|;Subs/Cont|Total||;Total|||",
            "164<-0,166<-1,167<-2,168<-3,169<-4,170<-5,171<-6,172<-7,173<-8,174<-9,175<-10,176<-11,177<-12,178<-13,179<-14,180<-15,181<-16,182<-17,183<-18,184<-19,185<-20,186<-21,187<-22,188<-23,189<-24,190<-25,191<-26,192<-27,193<-28,194<-29",
            [165], []),

        // ASE4 (BOGOTA LIMPIA): 282..285 sin fuente; ceros: canal 2 (`Vlr Intereses`) y grupo canal 9
        // en las dos mitades (Anticipos y Subs/Cont).
        new(HojaR2, 282, 328,
            "|||Vlr Servicio;|||Vlr Intereses;||A|Total;|EAB|Total|;|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENEL|Total|;|||Vlr Servicio;|||Vlr Intereses;||1|Total;|||Vlr Servicio;||2|Total;|||Vlr Servicio;|||Vlr Intereses;||5|Total;|NUEVO ESQUEMA|Total|;|||Vlr Servicio;|||Vlr Intereses;||H|Total;|||Vlr Servicio;|||Vlr Intereses;||O|Total;|||Vlr Servicio;|||Vlr Intereses;||T|Total;|OCCIDENTE|Total|;Componente|Total||;|||Subsidio(-)/Contribucion(+);||E|Total;|ENEL|Total|;|||Subsidio(-)/Contribucion(+);||1|Total;|||Subsidio(-)/Contribucion(+);||2|Total;|||Subsidio(-)/Contribucion(+);||5|Total;|NUEVO ESQUEMA|Total|;|||Subsidio(-)/Contribucion(+);||H|Total;|||Subsidio(-)/Contribucion(+);||O|Total;|||Subsidio(-)/Contribucion(+);||T|Total;|OCCIDENTE|Total|;Subs/Cont|Total||;Total|||",
            "|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENEL|Total|;|||Vlr Servicio;|||Vlr Intereses;||1|Total;|||Vlr Servicio;|||Vlr Intereses;||2|Total;|||Vlr Servicio;|||Vlr Intereses;||5|Total;|||Vlr Servicio;|||Vlr Intereses;||9|Total;|NUEVO ESQUEMA|Total|;|||Vlr Servicio;|||Vlr Intereses;||H|Total;|||Vlr Servicio;|||Vlr Intereses;||O|Total;|||Vlr Servicio;|||Vlr Intereses;||T|Total;|OCCIDENTE|Total|;Componente|Total||;|||Subsidio(-)/Contribucion(+);||E|Total;|ENEL|Total|;|||Subsidio(-)/Contribucion(+);||1|Total;|||Subsidio(-)/Contribucion(+);||2|Total;|||Subsidio(-)/Contribucion(+);||5|Total;|||Subsidio(-)/Contribucion(+);||9|Total;|NUEVO ESQUEMA|Total|;|||Subsidio(-)/Contribucion(+);||H|Total;|||Subsidio(-)/Contribucion(+);||O|Total;|||Subsidio(-)/Contribucion(+);||T|Total;|OCCIDENTE|Total|;Subs/Cont|Total||;Total|||",
            "286<-0,287<-1,288<-2,289<-3,290<-4,291<-5,292<-6,293<-7,294<-9,295<-10,296<-11,297<-12,298<-16,299<-17,300<-18,301<-19,302<-20,303<-21,304<-22,305<-23,306<-24,307<-25,308<-26,309<-27,310<-28,311<-29,312<-30,313<-31,314<-32,315<-33,316<-34,317<-35,318<-36,319<-39,320<-40,321<-41,322<-42,323<-43,324<-44,325<-45,326<-46,327<-47,328<-48",
            [282, 283, 284, 285],
            [
                (8, "|||Vlr Intereses", "2/NUEVO ESQUEMA", -8785.59m),
                (13, "|||Vlr Servicio", "9/NUEVO ESQUEMA", 20152773.11m),
                (14, "|||Vlr Intereses", "9/NUEVO ESQUEMA", -2377.86m),
                (15, "||9|Total", "9/NUEVO ESQUEMA", 20150395.25m),
                (37, "|||Subsidio(-)/Contribucion(+)", "9/NUEVO ESQUEMA", 5663564.75m),
                (38, "||9|Total", "9/NUEVO ESQUEMA", 5663564.75m),
            ]),

        // ASE5 (AREA LIMPIA): dos grupos E (ENEL y ENERBIT) → cada par toma su canal por cierre
        // (`R388 C=E`, `R392 C=E`); 382..385 sin fuente.
        new(HojaR2, 382, 419,
            "|||Vlr Servicio;|||Vlr Intereses;||A|Total;|EAB|Total|;|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENEL|Total|;|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENERBIT|Total|;|||Vlr Servicio;|||Vlr Intereses;||H|Total;|||Vlr Servicio;|||Vlr Intereses;||O|Total;|||Vlr Servicio;|||Vlr Intereses;||T|Total;|OCCIDENTE|Total|;Componente|Total||;|||Subsidio(-)/Contribucion(+);||E|Total;|ENEL|Total|;|||Subsidio(-)/Contribucion(+);||E|Total;|ENERBIT|Total|;|||Subsidio(-)/Contribucion(+);||H|Total;|||Subsidio(-)/Contribucion(+);||O|Total;|||Subsidio(-)/Contribucion(+);||T|Total;|OCCIDENTE|Total|;Subs/Cont|Total||;Total|||",
            "|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENEL|Total|;|||Vlr Servicio;|||Vlr Intereses;||E|Total;|ENERBIT|Total|;|||Vlr Servicio;|||Vlr Intereses;||H|Total;|||Vlr Servicio;|||Vlr Intereses;||O|Total;|||Vlr Servicio;|||Vlr Intereses;||T|Total;|OCCIDENTE|Total|;Componente|Total||;|||Subsidio(-)/Contribucion(+);||E|Total;|ENEL|Total|;|||Subsidio(-)/Contribucion(+);||E|Total;|ENERBIT|Total|;|||Subsidio(-)/Contribucion(+);||H|Total;|||Subsidio(-)/Contribucion(+);||O|Total;|||Subsidio(-)/Contribucion(+);||T|Total;|OCCIDENTE|Total|;Subs/Cont|Total||;Total|||",
            "386<-0,387<-1,388<-2,389<-3,390<-4,391<-5,392<-6,393<-7,394<-8,395<-9,396<-10,397<-11,398<-12,399<-13,400<-14,401<-15,402<-16,403<-17,404<-18,405<-19,406<-20,407<-21,408<-22,409<-23,410<-24,411<-25,412<-26,413<-27,414<-28,415<-29,416<-30,417<-31,418<-32,419<-33",
            [382, 383, 384, 385], []),
    ];

    // ── P6 (H1): matriz `Reversion Pagos R4` — bloques de BloquesR4 ──────────────────────────────

    private static readonly Bloque[] MatrizR4 =
    [
        new(HojaR4, 3, 15,
            "||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;ENEL|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;OCCIDENTE|Total|;Total||",
            "||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;ENEL|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;OCCIDENTE|Total|;Total||",
            "3<-0,4<-1,5<-2,6<-3,7<-4,8<-5,9<-6,10<-7,11<-8,12<-9,13<-10,14<-11,15<-12",
            [], []),

        // ASE2 (LIME): la fuente no trae el grupo NUEVO ESQUEMA (99..104) → el grupo OCCIDENTE de la
        // fuente (src#6..10) cae en 105..109 (su canal por cierre `R110 OCCIDENTE`), no en 99..103.
        new(HojaR4, 93, 111,
            "||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;ENEL|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;NUEVO ESQUEMA|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;OCCIDENTE|Total|;Total||",
            "||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;ENEL|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;OCCIDENTE|Total|;Total||",
            "93<-0,94<-1,95<-2,96<-3,97<-4,98<-5,105<-6,106<-7,107<-8,108<-9,109<-10,110<-11,111<-12",
            [99, 100, 101, 102, 103, 104], []),

        new(HojaR4, 194, 206,
            "||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;ENEL|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;OCCIDENTE|Total|;Total||",
            "||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;ENEL|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;OCCIDENTE|Total|;Total||",
            "194<-0,195<-1,196<-2,197<-3,198<-4,199<-5,200<-6,201<-7,202<-8,203<-9,204<-10,205<-11,206<-12",
            [], []),

        new(HojaR4, 231, 249,
            "||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;ENEL|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;NUEVO ESQUEMA|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;OCCIDENTE|Total|;Total||",
            "||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;ENEL|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;NUEVO ESQUEMA|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;OCCIDENTE|Total|;Total||",
            "231<-0,232<-1,233<-2,234<-3,235<-4,236<-5,237<-6,238<-7,239<-8,240<-9,241<-10,242<-11,243<-12,244<-13,245<-14,246<-15,247<-16,248<-17,249<-18",
            [], []),

        new(HojaR4, 346, 358,
            "||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;ENEL|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;OCCIDENTE|Total|;Total||",
            "||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;ENEL|Total|;||Vlr Servicio;||Vlr Intereses;|Componente|Total;||Subsidio(-)/Contribucion(+);|Subs/Cont|Total;OCCIDENTE|Total|;Total||",
            "346<-0,347<-1,348<-2,349<-3,350<-4,351<-5,352<-6,353<-7,354<-8,355<-9,356<-10,357<-11,358<-12",
            [], []),
    ];

    // ── S0/G-SHA: la plantilla oficial es intocable ──────────────────────────────────────────────

    [Fact]
    public void T0P47_Sha256_OficialIntacto()
    {
        using var stream = File.OpenRead(Oficial());
        Assert.Equal(Sha256Oficial, Convert.ToHexString(SHA256.HashData(stream)));
    }

    // ── P6: matriz de filas (etiqueta + canal → slot) congelada, por ASE ────────────────────────

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void T0P47_P6_R2_MatrizFilasYEmparejamiento_Congelada(int ase)
    {
        var bloque = MatrizR2[ase - 1];
        var oficial = Leer(Oficial(), bloque.Hoja);
        Assert.Equal(Separar(bloque.DestinoFirmas), FirmasDestino(oficial, bloque, ColsR2));

        var fuente = LeerFuenteDetalle(ase, esR2: true);
        Assert.Equal(Separar(bloque.FuenteFirmas), fuente.Filas.Select(f => f.Firma).ToArray());

        var pares = Emparejar(bloque, oficial, fuente, ColsR2, esR2: true);
        Assert.Equal(ParsearPares(bloque.Pares), pares);
        Assert.Equal(bloque.SinFuente, FilasSinFuente(bloque, pares));
        Assert.Equal(bloque.CerosDocumentados, CerosDelEmparejamiento(bloque, oficial, fuente, ColsR2, esR2: true));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void T0P47_P6_R4_MatrizFilasYEmparejamiento_Congelada(int ase)
    {
        var bloque = MatrizR4[ase - 1];
        var oficial = Leer(Oficial(), bloque.Hoja);
        Assert.Equal(Separar(bloque.DestinoFirmas), FirmasDestino(oficial, bloque, ColsR4));

        var fuente = LeerFuenteDetalle(ase, esR2: false);
        Assert.Equal(Separar(bloque.FuenteFirmas), fuente.Filas.Select(f => f.Firma).ToArray());

        var pares = Emparejar(bloque, oficial, fuente, ColsR4, esR2: false);
        Assert.Equal(ParsearPares(bloque.Pares), pares);
        Assert.Equal(bloque.SinFuente, FilasSinFuente(bloque, pares));
        Assert.Equal(bloque.CerosDocumentados, CerosDelEmparejamiento(bloque, oficial, fuente, ColsR4, esR2: false));
    }

    // ── P6: (encabezado de componente) → (columna destino) ──────────────────────────────────────

    [Fact]
    public void T0P47_P6_R2_ColumnasPorEncabezado_DestinoYFuente()
    {
        var oficial = Leer(Oficial(), HojaR2);
        Assert.Equal(
            [
                "Total", "Componente TDF", "Componente TTL", "Componente TVIAT", "Aprovechamiento",
                "CCSA Prest.Aprov.", "Especiales", "Componente TCS", "Componente TLU", "Componente TBL",
                "Componente TRT", "CCSA Prest. No Aprov.", "Deb/Cred", "Deb/Cred-otros",
            ],
            oficial.EncabezadosDesde(2, "E"));

        Assert.Equal(14, WorkbookLeafCellMapDetalleR2R4.ColumnasR2.Count);
        Assert.Equal("E", WorkbookLeafCellMapDetalleR2R4.ColumnasR2[0].Columna);
        Assert.Equal("Total", WorkbookLeafCellMapDetalleR2R4.ColumnasR2[0].Encabezado);
        Assert.Equal("R", WorkbookLeafCellMapDetalleR2R4.ColumnasR2[^1].Columna);
        Assert.Equal("Deb/Cred-otros", WorkbookLeafCellMapDetalleR2R4.ColumnasR2[^1].Encabezado);

        string[] sinEspeciales =
        [
            "Total", "Componente TDF", "Componente TTL", "Componente TVIAT", "Aprovechamiento",
            "CCSA Prest.Aprov.", "Componente TCS", "Componente TLU", "Componente TBL", "Componente TRT",
            "CCSA Prest. No Aprov.", "Deb/Cred",
        ];
        string[] conEspeciales =
        [
            "Total", "Componente TDF", "Componente TTL", "Componente TVIAT", "Aprovechamiento",
            "CCSA Prest.Aprov.", "Especiales", "Componente TCS", "Componente TLU", "Componente TBL",
            "Componente TRT", "CCSA Prest. No Aprov.", "Deb/Cred",
        ];
        foreach (var ase in new[] { 1, 2, 4 })
        {
            Assert.Equal(sinEspeciales, LeerFuenteDetalle(ase, esR2: true).Encabezados);
        }

        foreach (var ase in new[] { 3, 5 })
        {
            Assert.Equal(conEspeciales, LeerFuenteDetalle(ase, esR2: true).Encabezados);
        }
    }

    [Fact]
    public void T0P47_P6_R4_ColumnasPorEncabezado_DestinoYFuente()
    {
        var oficial = Leer(Oficial(), HojaR4);
        Assert.Equal(
            [
                "Total", "Componente TDF", "Componente TTL", "Componente TVIAT", "Aprovechamiento",
                "CCSA Prest.Aprov.", "Componente TCS", "Componente TLU", "Componente TBL", "Componente TRT",
                "CCSA Prest. No Aprov.", "Ajuste Decena", "SERVICIO ESPECIALES", "Ajuste Recargo Aseo", "Otros",
            ],
            oficial.EncabezadosDesde(2, "D"));

        string[] fuente12 =
        [
            "Total", "Componente TDF", "Componente TTL", "Componente TVIAT", "Aprovechamiento",
            "CCSA Prest.Aprov.", "Componente TCS", "Componente TLU", "Componente TBL", "Componente TRT",
            "CCSA Prest. No Aprov.", "Ajuste Decena",
        ];
        foreach (var ase in new[] { 1, 2, 3, 5 })
        {
            Assert.Equal(fuente12, LeerFuenteDetalle(ase, esR2: false).Encabezados);
        }

        Assert.Equal([.. fuente12, "Otros"], LeerFuenteDetalle(4, esR2: false).Encabezados);

        // Plan 47 (H1.3): `Otros → R` declarada por ENCABEZADO (13 columnas); las otras dos columnas
        // del oficial (`SERVICIO ESPECIALES P`, `Ajuste Recargo Aseo Q`) no tienen contraparte en
        // ninguna fuente julio-2/agosto-2 (ausencia probada) y siguen en el 0 de la plantilla.
        Assert.Equal(13, WorkbookLeafCellMapDetalleR2R4.ColumnasR4.Count);
        Assert.Equal("R", WorkbookLeafCellMapDetalleR2R4.ColumnasR4[^1].Columna);
        Assert.Equal("Otros", WorkbookLeafCellMapDetalleR2R4.ColumnasR4[^1].Encabezado);
        Assert.DoesNotContain(WorkbookLeafCellMapDetalleR2R4.ColumnasR4, c => c.Encabezado == "SERVICIO ESPECIALES");
        Assert.DoesNotContain(WorkbookLeafCellMapDetalleR2R4.ColumnasR4, c => c.Encabezado == "Ajuste Recargo Aseo");
    }

    // ── P2: canal efectivo (R46-1) — identidad de canal en TODO par + regla del predicado ───────

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void T0P47_P2_R2_TodoPar_ComparteCanalEfectivo(int ase)
    {
        var bloque = MatrizR2[ase - 1];
        var oficial = Leer(Oficial(), bloque.Hoja);
        var fuente = LeerFuenteDetalle(ase, esR2: true);
        var rotulosDestino = RotulosBloque(oficial, bloque);
        var rotulosFuente = fuente.Filas.Select(f => new Rotulo(f.A, f.B, f.C, f.D)).ToList();
        var pares = ParsearPares(bloque.Pares);

        foreach (var (fila, indice) in pares)
        {
            var canalFuente = CanalLocal(rotulosFuente, indice, PropiaR2, GrupoR2, SegundoR2);
            var canalDestino = CanalLocal(rotulosDestino, fila - bloque.Inicio, PropiaR2, GrupoR2, SegundoR2);
            Assert.True(
                string.Equals(canalFuente, canalDestino, StringComparison.OrdinalIgnoreCase),
                $"ASE{ase}: el par {fila}<-{indice} tiene canal fuente '{canalFuente}' vs canal slot '{canalDestino}'.");
        }

        // El grupo A/EAB sin fuente queda en el 0 de la plantilla: ASE1/2/4/5 (ASE3 sí lo trae).
        Assert.Equal(bloque.SinFuente, FilasSinFuente(bloque, pares));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void T0P47_P2_R4_TodoPar_ComparteCanalEfectivo(int ase)
    {
        var bloque = MatrizR4[ase - 1];
        var oficial = Leer(Oficial(), bloque.Hoja);
        var fuente = LeerFuenteDetalle(ase, esR2: false);
        var rotulosDestino = RotulosBloque(oficial, bloque);
        var rotulosFuente = fuente.Filas.Select(f => new Rotulo(f.A, f.B, f.C, string.Empty)).ToList();
        var pares = ParsearPares(bloque.Pares);

        foreach (var (fila, indice) in pares)
        {
            var canalFuente = CanalLocal(rotulosFuente, indice, PropiaR4, GrupoR4, SegundoR4);
            var canalDestino = CanalLocal(rotulosDestino, fila - bloque.Inicio, PropiaR4, GrupoR4, SegundoR4);
            Assert.True(
                string.Equals(canalFuente, canalDestino, StringComparison.OrdinalIgnoreCase),
                $"ASE{ase}: el par {fila}<-{indice} tiene canal fuente '{canalFuente}' vs canal slot '{canalDestino}'.");
        }

        Assert.Equal(bloque.SinFuente, FilasSinFuente(bloque, pares));
    }

    /// <summary>
    /// El espejo local del predicado (independiente) coincide con el predicado de PRODUCCIÓN
    /// (<see cref="CanalEfectivoDetalleR2R4"/>) en cada fila de cada bloque de R2/R4, para los dos
    /// períodos Q2 disponibles: si divergen, el fixture deja de describir el contrato vigente.
    /// </summary>
    [Theory]
    [InlineData("Prueba Julio-2")]
    [InlineData("Prueba Agosto-2")]
    public void T0P47_P2_PredicadoCore_Y_EspejoLocal_Coinciden(string periodo)
    {
        foreach (var ase in new[] { 1, 2, 3, 4, 5 })
        {
            var oficial = Leer(Oficial(), HojaR2);
            var bloque = MatrizR2[ase - 1];
            var destino = RotulosBloque(oficial, bloque);
            for (var i = 0; i < destino.Count; i++)
            {
                Assert.Equal(
                    CanalEfectivoDetalleR2R4.CanalEfectivo(destino, i, PropiaR2, GrupoR2, SegundoR2),
                    CanalLocal(destino, i, PropiaR2, GrupoR2, SegundoR2));
            }

            var fuente = LeerFuenteDetalle(ase, esR2: true, periodo);
            var fuenteRotulos = fuente.Filas.Select(f => new Rotulo(f.A, f.B, f.C, f.D)).ToList();
            for (var i = 0; i < fuenteRotulos.Count; i++)
            {
                Assert.Equal(
                    CanalEfectivoDetalleR2R4.CanalEfectivo(fuenteRotulos, i, PropiaR2, GrupoR2, SegundoR2),
                    CanalLocal(fuenteRotulos, i, PropiaR2, GrupoR2, SegundoR2));
            }

            var oficialR4 = Leer(Oficial(), HojaR4);
            var bloqueR4 = MatrizR4[ase - 1];
            var destinoR4 = RotulosBloque(oficialR4, bloqueR4);
            for (var i = 0; i < destinoR4.Count; i++)
            {
                Assert.Equal(
                    CanalEfectivoDetalleR2R4.CanalEfectivo(destinoR4, i, PropiaR4, GrupoR4, SegundoR4),
                    CanalLocal(destinoR4, i, PropiaR4, GrupoR4, SegundoR4));
            }

            var fuenteR4 = LeerFuenteDetalle(ase, esR2: false, periodo);
            var fuenteR4Rotulos = fuenteR4.Filas.Select(f => new Rotulo(f.A, f.B, f.C, string.Empty)).ToList();
            for (var i = 0; i < fuenteR4Rotulos.Count; i++)
            {
                Assert.Equal(
                    CanalEfectivoDetalleR2R4.CanalEfectivo(fuenteR4Rotulos, i, PropiaR4, GrupoR4, SegundoR4),
                    CanalLocal(fuenteR4Rotulos, i, PropiaR4, GrupoR4, SegundoR4));
            }
        }
    }

    /// <summary>
    /// W-3 (la vía fuerte): con DOS grupos del mismo canal en el bloque (R2 ASE5: <c>E</c> de ENEL y
    /// <c>E</c> de ENERBIT), el canal compuesto por segundo cierre desambigua por GEOMETRÍA. Fuente
    /// SINTÉTICA (los mismos bytes reales, con los dos grupos en orden INVERTIDO) → cada grupo recibe
    /// su propio slot (386..389 para ENEL, 390..393 para ENERBIT): el resultado NO depende del orden.
    /// Sin el compuesto ambos grupos comparten canal (<c>E</c>) y el empate se resolvería por posición.
    /// </summary>
    [Fact]
    public void T0P47_W3_R2_Ase5_DobleGrupoE_OrdenInvertido_CadaGrupoRecibeLoSuyo()
    {
        var bloque = MatrizR2[4];
        var oficial = Leer(Oficial(), HojaR2);
        var real = LeerFuenteDetalle(5, esR2: true);
        var filas = real.Filas.ToList();

        // Estructura real: [0..3] = grupo E/ENEL (VS, VI, `||E|Total`, `|ENEL|Total|`) y
        // [4..7] = grupo E/ENERBIT. Se reordena SOLO la secuencia observada (bytes intactos).
        var invertida = new List<FilaFuente>();
        invertida.AddRange(filas.Skip(4).Take(4));
        invertida.AddRange(filas.Take(4));
        invertida.AddRange(filas.Skip(8));
        var fuenteInvertida = new Fuente(invertida, real.Encabezados);
        var rotulosInvertidos = invertida.Select(f => new Rotulo(f.A, f.B, f.C, f.D)).ToList();

        // Cada grupo conserva la clave de su empresa (canal compuesto).
        Assert.Equal("E/ENERBIT", CanalLocal(rotulosInvertidos, 0, PropiaR2, GrupoR2, SegundoR2));
        Assert.Equal("E/ENEL", CanalLocal(rotulosInvertidos, 4, PropiaR2, GrupoR2, SegundoR2));
        Assert.Equal(7522088.94m, invertida[0].Total);   // ENERBIT `Vlr Servicio`
        Assert.Equal(19713558.23m, invertida[4].Total);  // ENEL `Vlr Servicio`

        var pares = Emparejar(bloque, oficial, fuenteInvertida, ColsR2, esR2: true);
        // ENERBIT (índices 0..3 tras invertir) → su grupo en el bloque (390..393).
        Assert.Equal(390, pares.Single(p => p.Value == 0).Key);
        Assert.Equal(391, pares.Single(p => p.Value == 1).Key);
        Assert.Equal(392, pares.Single(p => p.Value == 2).Key);
        Assert.Equal(393, pares.Single(p => p.Value == 3).Key);
        // ENEL (índices 4..7 tras invertir) → el suyo (386..389), con SU valor.
        Assert.Equal(386, pares.Single(p => p.Value == 4).Key);
        Assert.Equal(387, pares.Single(p => p.Value == 5).Key);
        Assert.Equal(388, pares.Single(p => p.Value == 6).Key);
        Assert.Equal(389, pares.Single(p => p.Value == 7).Key);

        // El resto del bloque sigue cayendo en su sitio (el reordenamiento no contagia).
        Assert.Equal(403, pares.Single(p => p.Value == 17).Key);  // OCCIDENTE (cierre del grupo T)

        // Prueba de que el compuesto es lo que lo permite: con canal simple, ENEL y ENERBIT comparten
        // rótulo (`E`) y quedarían indistinguibles (empate por orden).
        Assert.Equal("E", CanalLocal(rotulosInvertidos, 0, PropiaR2, GrupoR2));
        Assert.Equal("E", CanalLocal(rotulosInvertidos, 4, PropiaR2, GrupoR2));

        // Y el predicado de PRODUCCIÓN coincide con el espejo en esta forma invertida.
        Assert.Equal(
            CanalEfectivoDetalleR2R4.CanalEfectivo(rotulosInvertidos, 4, PropiaR2, GrupoR2, SegundoR2),
            CanalLocal(rotulosInvertidos, 4, PropiaR2, GrupoR2, SegundoR2));
    }

    // ── P4: ceros documentados (sin slot probado) — lista Core ≡ espejo, en los dos períodos ────

    /// <summary>
    /// Plan 48 (H1.1 — T0 §2.2): la geometría v2 tiene slot para el canal 9 y el `Vlr Intereses` del
    /// canal 2 de ASE2/ASE4. Las 9 filas de julio-2 del Plan 47 eran un **falso hueco** y se retiraron
    /// de `CerosDocumentados`; solo sobreviven las 3 claves A/EAB de agosto-2 (ausencia probada en la
    /// mitad Subs/Cont de ASE4 v2). El fixture v1 histórico conserva sus matrices; este test certifica
    /// el contrato NUEVO contra los bytes de la v2.
    /// </summary>
    [Fact]
    public void T0P48_H1_CerosDocumentados_SoloClavesAusentesEnV2()
    {
        Assert.Equal(3, CerosDocumentadosDetalleR2R4.TotalFilasDocumentadas);
        Assert.Equal(3, CerosDocumentadosDetalleR2R4.Entradas.Count);
        Assert.All(CerosDocumentadosDetalleR2R4.Entradas, e => Assert.Equal(HojaDetalleR2R4.R2, e.Hoja));
        Assert.All(CerosDocumentadosDetalleR2R4.Entradas, e => Assert.Equal(4, e.Ase));
        Assert.All(CerosDocumentadosDetalleR2R4.Entradas, e => Assert.True(e.Canal is null || e.Canal.Contains("A/EAB")));

        // Las claves del falso hueco ya NO se documentan (si no hubiera slot serían ERR-SIN-SLOT).
        Assert.Equal(0, CerosDocumentadosDetalleR2R4.CantidadDocumentada(HojaDetalleR2R4.R2, 2, "|||Vlr Servicio", "9/NUEVO ESQUEMA"));
        Assert.Equal(0, CerosDocumentadosDetalleR2R4.CantidadDocumentada(HojaDetalleR2R4.R2, 2, "|||Vlr Intereses", "2/NUEVO ESQUEMA"));
        Assert.Equal(0, CerosDocumentadosDetalleR2R4.CantidadDocumentada(HojaDetalleR2R4.R2, 2, "||9|Total", "9/NUEVO ESQUEMA"));
        Assert.Equal(0, CerosDocumentadosDetalleR2R4.CantidadDocumentada(HojaDetalleR2R4.R2, 4, "|||Vlr Servicio", "9/NUEVO ESQUEMA"));
        Assert.Equal(0, CerosDocumentadosDetalleR2R4.CantidadDocumentada(HojaDetalleR2R4.R2, 4, "|||Vlr Intereses", "9/NUEVO ESQUEMA"));
        Assert.Equal(0, CerosDocumentadosDetalleR2R4.CantidadDocumentada(HojaDetalleR2R4.R2, 4, "||9|Total", "9/NUEVO ESQUEMA"));
        Assert.Equal(0, CerosDocumentadosDetalleR2R4.CantidadDocumentada(HojaDetalleR2R4.R2, 4, "|||Subsidio(-)/Contribucion(+)", "9/NUEVO ESQUEMA"));
    }

    [Fact]
    public void T0P48_H1_V2_TieneLosSlotsDelFalsoHueco()
    {
        var v2 = Leer(V2(), HojaR2);
        Assert.Equal(9m, v2.Numero("C89"));
        Assert.Equal("NUEVO ESQUEMA", v2.Texto("B", 90));
        Assert.Equal(9m, v2.Numero("C315"));
        Assert.Equal("NUEVO ESQUEMA", v2.Texto("B", 316));
        Assert.Equal(3m, v2.Numero("C83"));
        Assert.Equal(3m, v2.Numero("C309"));

        // Las 3 claves A/EAB sobrevivientes siguen ausentes: ASE4 v2 tiene UN solo grupo A/EAB.
        Assert.Single(Enumerable.Range(293, 58), n => v2.Texto("C", n) == "A");
        Assert.Single(Enumerable.Range(293, 58), n => v2.Texto("B", n) == "EAB");
    }

    // ── P4 (H1.2): el falso hueco quedó corregido — los slots del canal E reciben su valor ───────

    [Fact]
    public void T0P47_P4_R2_FalsoHuecoCorregido_SlotsCanalE()
    {
        // ASE1 (PROMOAMBIENTAL): la fuente no trae el grupo A/EAB. El `Vlr Servicio`/`Vlr Intereses` del
        // canal E (cerrados por `R9 C=E D=Total` en el bloque) van a las filas 7/8; 3..6 quedan en el 0
        // de la plantilla. Antes de H1 esos dos valores caían en 3/4 (mapeo por orden).
        var bloque = MatrizR2[0];
        var pares = ParsearPares(bloque.Pares);
        Assert.Equal(0, pares[7]);
        Assert.Equal(1, pares[8]);
        Assert.DoesNotContain(3, pares.Keys);
        Assert.DoesNotContain(4, pares.Keys);
        Assert.Equal([3, 4, 5, 6], bloque.SinFuente);

        var oficial = Leer(Oficial(), HojaR2);
        Assert.Equal("Vlr Servicio", oficial.Texto("D", 7));
        Assert.Equal("0", oficial.Texto("E", 7));      // el OFICIAL sigue en 0: lo escribe la corrida
        Assert.False(oficial.TieneFormula("E7"));
        Assert.Equal("A", oficial.Texto("C", 5));      // cierre del grupo A (3..4)
        Assert.Equal("E", oficial.Texto("C", 9));      // cierre del grupo E (7..8)

        var fuente = LeerFuenteDetalle(1, esR2: true);
        Assert.Equal("E/ENEL", CanalLocal(fuente.Filas.Select(f => new Rotulo(f.A, f.B, f.C, f.D)).ToList(), 0, PropiaR2, GrupoR2, SegundoR2));
        Assert.Equal(104754634.94m, fuente.Filas[0].Total);
        Assert.Equal(15816.77m, fuente.Filas[1].Total);
    }

    [Fact]
    public void T0P47_P4_R4_Ase4_Otros_DeclaradaPorEncabezado()
    {
        // Plan 47 (H1.3): la fuente R4 de ASE4 trae `Otros` (13.º encabezado) y el oficial tiene la
        // columna R `Otros` como VALOR (literal 0, sin fórmula) → el mapa la declara y el writer escribe
        // el dato (p. ej. -0,29 en `Vlr Intereses`), en vez de descartarlo (falso hueco de columna).
        var fuente = LeerFuenteDetalle(4, esR2: false);
        Assert.Contains("Otros", fuente.Encabezados);
        Assert.Equal(-0.29m, fuente.Filas[1].Valor("Otros"));

        var oficial = Leer(Oficial(), HojaR4);
        Assert.Equal("Otros", oficial.Texto("R", 2));
        Assert.Equal("0", oficial.Texto("R", 232));
        Assert.False(oficial.TieneFormula("R232"));
        Assert.Contains(WorkbookLeafCellMapDetalleR2R4.ColumnasR4, c => c is { Encabezado: "Otros", Columna: "R" });
    }

    // ── P1/P4: SALDOS POR NOTA — fuente propia, mapeo por título (sin orden) ────────────────────

    [Fact]
    public void T0P47_SaldosPorNota_MapeoPorTitulo_SinEmparejamientoPorOrden()
    {
        var oficial = Leer(Oficial(), "SALDOS POR NOTA");
        Assert.Equal(
            [
                "Total", "Componente TDF", "Componente TTL", "Componente TVIAT", "Aprovechamiento",
                "CCSA Prest.Aprov.", "Especiales", "Componente TCS", "Componente TLU", "Componente TBL",
                "Componente TRT", "CCSA Prest. No Aprov.", "Deb/Cred",
            ],
            oficial.EncabezadosDesde(2, "C"));

        Assert.Equal(
            [
                ("Vlr Servicio", ""), ("Vlr Intereses", ""), ("Total", "Componente"),
                ("Subsidio(-)/Contribucion(+)", ""), ("Total", "Subs/Cont"), ("", "Total"),
            ],
            Enumerable.Range(3, 6).Select(n => (oficial.Texto("B", n), oficial.Texto("A", n))).ToArray());

        var fuente = Leer(RutaSaldosNotas(1), "Sheet1");
        var encabezados = fuente.EncabezadosDesde(3, "C");
        Assert.DoesNotContain("Especiales", encabezados);
        Assert.Equal("Componente TCS", encabezados[6]);
        Assert.Equal(796773.47m, fuente.Numero("C4"));
        Assert.Equal("Vlr Servicio", fuente.Texto("B", 4));
        Assert.Equal(
            ["Vlr Servicio", "Componente", "Subsidio(-)/Contribucion(+)", "Subs/Cont", "Total"],
            Enumerable.Range(4, 5)
                .Select(n =>
                {
                    var b = fuente.Texto("B", n);
                    return b.Length == 0 || b == "Total" ? fuente.Texto("A", n) : b;
                })
                .ToArray());
    }

    // ── P1/P5: `Recaudo *` es pegado real desde Conciliaciones (RESUMEN MES) ────────────────────

    [Fact]
    public void T0P47_Recaudo_EsPegadoReal_DesdeResumenMes_PosicionalHoyAlineado()
    {
        var oficial = Leer(Oficial(), "Recaudo ENEL");
        Assert.Equal("ASE", oficial.Texto("C", 2));
        Assert.Equal("VALOR 2° Q", oficial.Texto("F", 2));
        Assert.Equal("NUMERO REG. 2°Q", oficial.Texto("G", 2));
        Assert.Equal(["1", "2", "3", "4", "5"], Enumerable.Range(3, 5).Select(n => oficial.Texto("C", n)).ToArray());

        var resumen = Leer(RutaConciliacion("Conjunta ENEL"), "RESUMEN MES");
        Assert.Equal("OPORTUNO", resumen.Texto("B", 2));
        Assert.Equal(["1", "2", "3", "4", "5"], Enumerable.Range(3, 5).Select(n => resumen.Texto("C", n)).ToArray());
        Assert.Equal(16467129562m, resumen.Numero("F3"));
        Assert.Equal(19037496395m, resumen.Numero("F4"));
        Assert.Equal(13952517350m, resumen.Numero("F5"));
        Assert.Equal(6579568126m, resumen.Numero("F6"));
        Assert.Equal(11684535987m, resumen.Numero("F7"));
        Assert.Equal(67721247420m, resumen.Numero("F9"));
        Assert.Equal(
            resumen.Numero("F9"),
            resumen.Numero("F3") + resumen.Numero("F4") + resumen.Numero("F5") + resumen.Numero("F6") + resumen.Numero("F7"));

        var otros = Leer(RutaConciliacion("Conjunta Otros"), "RESUMEN MES");
        Assert.All(Enumerable.Range(3, 5), n => Assert.Equal(0m, otros.Numero("F" + n.ToString(CultureInfo.InvariantCulture))));
        Assert.Equal(0m, otros.Numero("F9"));
    }

    // ── P5: hojas de período (sello septiembre) = PLACEHOLDER que el preparador reescribe ───────

    [Fact]
    public void T0P47_HojasDePeriodo_SonPlaceholderDeSeptiembre()
    {
        var nombres = NombresDeHoja(Oficial());
        Assert.Contains("DetRetri2026092", nombres);
        Assert.Contains("DetValiRetri2026092", nombres);
        Assert.Contains("Informe AFaseo Recaudo 202609-2", nombres);
        Assert.DoesNotContain("DetRetri2026072", nombres);
        Assert.DoesNotContain("DetValiRetri2026072", nombres);
        Assert.DoesNotContain("Informe AFaseo Recaudo 202607-2", nombres);

        var consolidado = Leer(Oficial(), "CONSOLIDADO_TOTAL RECAUDO");
        Assert.Equal("2026092", consolidado.Texto("N", 3));
        Assert.False(consolidado.TieneFormula("N3"));
        Assert.Equal(new DateTime(2026, 9, 16), DateTime.FromOADate((double)consolidado.Numero("G7")));
        Assert.Equal(new DateTime(2026, 9, 30), DateTime.FromOADate((double)consolidado.Numero("K7")));

        var ajustes = Leer(Oficial(), "AJUSTES - SF-T");
        Assert.True(ajustes.TieneFormula("N3"));
        Assert.Contains("CONSOLIDADO_TOTAL RECAUDO", ajustes.Formula("N3"));
        var bce = Leer(Oficial(), "BCE SC POR FACT.");
        Assert.Contains("DetRetri2026092", bce.Formula("H3"));
    }

    // ── P1: tabla dentro/fuera con criterio citado por hoja ────────────────────────────────────

    [Fact]
    public void T0P47_P1_TablaDentroFuera_CriteriosCitados()
    {
        var consolidado = Leer(Oficial(), "CONSOLIDADO_TOTAL RECAUDO");
        var ajustes = Leer(Oficial(), "AJUSTES - SF-T");
        var validacion = Leer(Oficial(), "VALIDACION_TOTAL");
        var banco = Leer(Oficial(), "REPORTE RECAUDO x BANCO");
        var bce = Leer(Oficial(), "BCE SC POR FACT.");
        var interventoria = Leer(Oficial(), "INTERVENTORIA");
        var validaControl = Leer(Oficial(), "Valida - Control Recaudo");

        Assert.Contains("Rem. Anticipos R2", consolidado.Formula("D28"));
        Assert.Contains("Rem. Anticipos R2", consolidado.Formula("D32"));
        Assert.Contains("SUM", consolidado.Formula("D33"));
        Assert.Equal("D9+D28", ajustes.Formula("D47"));
        Assert.Equal("SUM(C3:C8)", validacion.Formula("C9"));
        Assert.Contains("Valida - Control Recaudo", validacion.Formula("C14"));
        Assert.Equal("C14=C9", validacion.Formula("C15"));
        Assert.True(validaControl.TieneFormula("C2") || validaControl.Texto("C2").Length > 0);
        Assert.Equal("Resumen Recaudo Aplicado Por Servicio", banco.Texto("B", 11));
        Assert.Equal(
            ["ENEL", "CIUDAD LIMPIA", "EAB", "NUEVO ESQUEMA", "ENERBIT", "OCCIDENTE", "Total"],
            Enumerable.Range(0, 7).Select(i => banco.Texto(LetraDe(2 + i), 12)).ToArray());
        Assert.Equal("1-APLICADOS A FACTURACION", banco.Texto("B", 13));
        Assert.Equal("7-APLICADOS A RECIBOS SERVICIOS ESPECIALES", banco.Texto("B", 16));
        Assert.Equal("SUBSIDIO", bce.Texto("D", 2));
        Assert.Equal("CONTRIBUCION", bce.Texto("E", 2));
        Assert.Equal(["1", "2", "3", "4", "5"], Enumerable.Range(3, 5).Select(n => bce.Texto("C", n)).ToArray());
        Assert.True(bce.TieneFormula("H3"));
        Assert.Contains("SUM", interventoria.Formula("K31"));
        Assert.Equal(1, Leer(RutaRetribucionNegativa(1), "Sheet1").MaxRow);
    }

    // ── Impacto: el corrimiento no altera la aritmética (ninguna <f> consume filas sin etiqueta) ─

    [Theory]
    [InlineData(HojaR2)]
    [InlineData(HojaR4)]
    public void T0P47_Impacto_NingunaFormulaConsumeFilasSinEtiqueta(string hoja)
    {
        var esR2 = hoja == HojaR2;
        var oficial = Leer(Oficial(), hoja);

        var sinEtiqueta = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var bloque in esR2 ? MatrizR2 : MatrizR4)
        {
            for (var fila = bloque.Inicio; fila <= bloque.Fin; fila++)
            {
                if (oficial.Texto("A", fila).Length != 0 || oficial.Texto("B", fila).Length != 0)
                {
                    continue;
                }

                for (var letra = 'A'; letra <= 'R'; letra++)
                {
                    sinEtiqueta.Add($"{letra}{fila}");
                }
            }
        }

        var refs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (_, formula) in oficial.Formulas)
        {
            var local = Regex.Replace(formula, @"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_. ]*)![A-Z]{1,3}\$?\d+", string.Empty);
            foreach (Match m in Regex.Matches(local, @"[A-Z]{1,3}\$?\d+"))
            {
                refs.Add(m.Value.Replace("$", string.Empty));
            }
        }

        Assert.NotEmpty(refs);
        Assert.NotEmpty(sinEtiqueta);
        Assert.DoesNotContain(sinEtiqueta, refs.Contains);
    }

    // ── Derivación (espejo local de las reglas de producción; SIN producción) ───────────────────

    private sealed record FilaFuente(
        int Indice,
        string A,
        string B,
        string C,
        string D,
        string Clave,
        IReadOnlyDictionary<string, decimal?> Valores)
    {
        public string Firma => Clave;

        public decimal Total => Valor("Total") ?? 0m;

        public decimal? Valor(string encabezado) =>
            Valores.TryGetValue(encabezado, out var valor) ? valor : null;
    }

    private sealed record Fuente(IReadOnlyList<FilaFuente> Filas, string[] Encabezados);

    /// <summary>Columnas de etiqueta propia / grupo / segundo cierre (0=A..3=D), espejo de los datos
    /// del mapa (`WorkbookLeafCellMapDetalleR2R4`).</summary>
    private static readonly int[] PropiaR2 = [0, 1, 2];
    private static readonly int[] GrupoR2 = [2];
    private static readonly int[] SegundoR2 = [1];
    private static readonly int[] PropiaR4 = [0];
    private static readonly int[] GrupoR4 = [0];
    private static readonly int[] SegundoR4 = [];

    private static string[] Separar(string firmas) => firmas.Split(';');

    /// <summary>Espejo LOCAL del predicado de canal efectivo (R46-1 + W-3): la etiqueta de grupo de la
    /// fila <c>Total</c> que cierra el grupo, COMPUESTA con la empresa del primer cierre de empresa por
    /// debajo cuando hay columnas de segundo cierre (R2); <c>null</c> si la fila trae etiqueta propia o
    /// no hay cierre. Se contrasta contra el predicado de producción en
    /// <c>T0P47_P2_PredicadoCore_Y_EspejoLocal_Coinciden</c>.</summary>
    private static string? CanalLocal(IReadOnlyList<Rotulo> filas, int indice, int[] propia, int[] grupo, int[]? segundo = null)
    {
        var fila = filas[indice];
        if (TieneEtiquetaPropia(fila, propia) && EtiquetaLocal(fila, grupo) is null)
        {
            return null;
        }

        var canal = EtiquetaLocal(fila, grupo);
        var desde = indice + 1;
        if (canal is null)
        {
            var cierre = -1;
            for (var j = indice + 1; j < filas.Count && cierre < 0; j++)
            {
                if (EsCierreLocal(filas[j]) && EtiquetaLocal(filas[j], grupo) is not null)
                {
                    cierre = j;
                }
            }

            if (cierre < 0)
            {
                return null;
            }

            canal = EtiquetaLocal(filas[cierre], grupo);
            desde = cierre + 1;
        }

        var empresa = segundo is { Length: > 0 } ? PrimerCierreConEtiquetaLocal(filas, desde, segundo) : null;
        return empresa is null ? canal : $"{canal}/{empresa}";
    }

    private static string? PrimerCierreConEtiquetaLocal(IReadOnlyList<Rotulo> filas, int desde, int[] columnas)
    {
        for (var j = desde; j < filas.Count; j++)
        {
            if (!EsCierreLocal(filas[j]))
            {
                continue;
            }

            var etiqueta = EtiquetaLocal(filas[j], columnas);
            if (etiqueta is not null)
            {
                return etiqueta;
            }
        }

        return null;
    }

    private static bool EsCierreLocal(Rotulo fila) =>
        new[] { fila.A, fila.B, fila.C, fila.D }.Any(t => t.Trim().Equals("Total", StringComparison.OrdinalIgnoreCase));

    private static string? EtiquetaLocal(Rotulo fila, int[] columnas)
    {
        foreach (var columna in columnas)
        {
            var texto = TextoDe(fila, columna).Trim();
            if (texto.Length > 0 && !texto.Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                return texto;
            }
        }

        return null;
    }

    private static bool TieneEtiquetaPropia(Rotulo fila, int[] propia) =>
        propia.Any(c =>
        {
            var texto = TextoDe(fila, c).Trim();
            return texto.Length > 0 && !texto.Equals("Total", StringComparison.OrdinalIgnoreCase);
        });

    private static string TextoDe(Rotulo fila, int columna) => fila.Texto(columna);

    private static string[] FirmasDestino(Hoja oficial, Bloque bloque, string[] columnas) =>
        Enumerable.Range(bloque.Inicio, bloque.Fin - bloque.Inicio + 1)
            .Select(n => string.Join("|", columnas.Select(c => oficial.Texto(c, n).Trim())))
            .ToArray();

    private static List<Rotulo> RotulosBloque(Hoja oficial, Bloque bloque) =>
        Enumerable.Range(bloque.Inicio, bloque.Fin - bloque.Inicio + 1)
            .Select(n => new Rotulo(oficial.Texto("A", n), oficial.Texto("B", n), oficial.Texto("C", n), oficial.Texto("D", n)))
            .ToList();

    /// <summary>Espejo LOCAL del emparejamiento de producción: clave <c>(firma + canal efectivo)</c>,
    /// primer slot libre; devuelve el mapa fila-destino absoluta → índice-fuente.</summary>
    private static Dictionary<int, int> Emparejar(Bloque bloque, Hoja oficial, Fuente fuente, string[] columnas, bool esR2)
    {
        var destino = RotulosBloque(oficial, bloque);
        var firmasDestino = destino.Select(r => FirmaDe(r, columnas)).ToList();
        var canalesDestino = Enumerable.Range(0, destino.Count)
            .Select(i => CanalLocal(destino, i, esR2 ? PropiaR2 : PropiaR4, esR2 ? GrupoR2 : GrupoR4, esR2 ? SegundoR2 : SegundoR4))
            .ToList();
        var fuenteRotulos = fuente.Filas.Select(f => new Rotulo(f.A, f.B, f.C, f.D)).ToList();
        var consumido = new bool[destino.Count];
        var pares = new Dictionary<int, int>();

        for (var i = 0; i < fuenteRotulos.Count; i++)
        {
            var firma = FirmaDe(fuenteRotulos[i], columnas);
            var canal = CanalLocal(fuenteRotulos, i, esR2 ? PropiaR2 : PropiaR4, esR2 ? GrupoR2 : GrupoR4, esR2 ? SegundoR2 : SegundoR4);

            for (var j = 0; j < destino.Count; j++)
            {
                if (consumido[j])
                {
                    continue;
                }

                if (!string.Equals(firmasDestino[j], firma, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.Equals(canalesDestino[j], canal, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                consumido[j] = true;
                pares[bloque.Inicio + j] = i;
                break;
            }
        }

        return pares;
    }

    /// <summary>Filas-fuente sin slot (ceros documentados) con su Total, en orden.</summary>
    private static (int Indice, string Firma, string? Canal, decimal Total)[] CerosDelEmparejamiento(
        Bloque bloque,
        Hoja oficial,
        Fuente fuente,
        string[] columnas,
        bool esR2,
        string? periodo = null)
    {
        _ = periodo;
        var pares = Emparejar(bloque, oficial, fuente, columnas, esR2);
        var fuenteRotulos = fuente.Filas.Select(f => new Rotulo(f.A, f.B, f.C, f.D)).ToList();
        var asignados = pares.Values.ToHashSet();

        return Enumerable.Range(0, fuenteRotulos.Count)
            .Where(i => !asignados.Contains(i))
            .Select(i => (
                i,
                FirmaDe(fuenteRotulos[i], columnas),
                CanalLocal(fuenteRotulos, i, esR2 ? PropiaR2 : PropiaR4, esR2 ? GrupoR2 : GrupoR4, esR2 ? SegundoR2 : SegundoR4),
                fuente.Filas[i].Total))
            .ToArray();
    }

    private static string FirmaDe(Rotulo rotulo, IReadOnlyList<string> columnas)
    {
        var partes = columnas.Select(c => rotulo.Texto(IndiceDeLetra(c)).Trim()).ToArray();
        return string.Join("|", partes);
    }

    private static int[] FilasSinFuente(Bloque bloque, Dictionary<int, int> pares) =>
        Enumerable.Range(bloque.Inicio, bloque.Fin - bloque.Inicio + 1)
            .Where(fila => !pares.ContainsKey(fila))
            .ToArray();

    private static Dictionary<int, int> ParsearPares(string pares)
    {
        var mapa = new Dictionary<int, int>();
        foreach (var par in pares.Split(',', StringSplitOptions.RemoveEmptyEntries))
        {
            var partes = par.Split("<-");
            mapa[int.Parse(partes[0], CultureInfo.InvariantCulture)] = int.Parse(partes[1], CultureInfo.InvariantCulture);
        }

        return mapa;
    }

    /// <summary>Filas (por firma) entre el cierre de un grupo y el cierre del grupo siguiente.</summary>
    private static string[] FilasEntreCierres(IReadOnlyList<string> firmas, string cierreInicial, string cierreFinal)
    {
        var lista = firmas.ToList();
        var inicio = lista.FindIndex(f => f.EndsWith($"|{cierreInicial}|Total", StringComparison.OrdinalIgnoreCase));
        var fin = lista.FindIndex(f => f.EndsWith($"|{cierreFinal}|Total", StringComparison.OrdinalIgnoreCase));
        Assert.True(inicio >= 0 && fin > inicio, $"No se ubicaron los cierres '{cierreInicial}'/'{cierreFinal}'.");
        return lista.Skip(inicio + 1).Take(fin - inicio - 1).ToArray();
    }

    // ── Lectura (BCL ZIP+XML, read-only) ────────────────────────────────────────────────────────

    private sealed class Hoja
    {
        private readonly Dictionary<string, string> _textos = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, decimal> _numeros = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _formulas = new(StringComparer.OrdinalIgnoreCase);

        public int MaxRow { get; private set; }

        public IEnumerable<(string Celda, string Formula)> Formulas =>
            _formulas.Where(f => f.Value.Length > 0).Select(f => (f.Key, f.Value));

        public void Agregar(string referencia, string? texto, decimal? numero, bool tieneFormula, string? formula)
        {
            if (texto is not null)
            {
                _textos[referencia] = texto;
            }

            if (numero is not null)
            {
                _numeros[referencia] = numero.Value;
            }

            if (tieneFormula)
            {
                _formulas[referencia] = formula ?? string.Empty;
            }

            var fila = int.Parse(new string(referencia.Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture);
            if (fila > MaxRow)
            {
                MaxRow = fila;
            }
        }

        public string Texto(string columna, int fila) => Texto($"{columna}{fila}");

        public string Texto(string referencia) =>
            _textos.TryGetValue(referencia, out var valor) ? valor : string.Empty;

        public decimal Numero(string referencia) => _numeros.TryGetValue(referencia, out var valor) ? valor : 0m;

        public decimal? NumeroOrNull(string referencia) =>
            _numeros.TryGetValue(referencia, out var valor) ? valor : null;

        public bool TieneFormula(string referencia) => _formulas.ContainsKey(referencia);

        public string Formula(string referencia) =>
            _formulas.TryGetValue(referencia, out var valor) ? valor : string.Empty;

        public string[] EncabezadosDesde(int fila, string columnaInicial)
        {
            var textos = new List<string>();
            for (var i = IndiceDeLetra(columnaInicial); i <= IndiceDeLetra("R"); i++)
            {
                textos.Add(Texto(LetraDe(i), fila));
            }

            while (textos.Count > 0 && textos[^1].Length == 0)
            {
                textos.RemoveAt(textos.Count - 1);
            }

            return textos.ToArray();
        }
    }

    private static Hoja Leer(string ruta, string nombreHoja)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var libro = Cargar(zip, "xl/workbook.xml");
        var sheet = libro.Root!.Element(Ns + "sheets")!.Elements(Ns + "sheet")
            .First(s => string.Equals(s.Attribute("name")?.Value, nombreHoja, StringComparison.Ordinal));
        var id = sheet.Attribute(RelNs + "id")!.Value;
        var rels = Cargar(zip, "xl/_rels/workbook.xml.rels");
        var destino = rels.Root!.Elements().First(r => r.Attribute("Id")?.Value == id).Attribute("Target")!.Value;
        if (!destino.StartsWith("xl/", StringComparison.Ordinal))
        {
            destino = "xl/" + destino.TrimStart('/');
        }

        var compartidas = new List<string>();
        if (zip.GetEntry("xl/sharedStrings.xml") is not null)
        {
            compartidas = Cargar(zip, "xl/sharedStrings.xml").Root!.Elements(Ns + "si").Select(si => si.Value).ToList();
        }

        var doc = Cargar(zip, destino);
        var hoja = new Hoja();
        foreach (var row in doc.Root!.Element(Ns + "sheetData")!.Elements(Ns + "row"))
        {
            foreach (var c in row.Elements(Ns + "c"))
            {
                var referencia = c.Attribute("r")?.Value;
                if (referencia is null)
                {
                    continue;
                }

                var tipo = c.Attribute("t")?.Value;
                var v = c.Element(Ns + "v")?.Value;
                string? texto = null;
                decimal? numero = null;
                if (tipo == "s" && v is not null && int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx)
                    && idx >= 0 && idx < compartidas.Count)
                {
                    texto = compartidas[idx];
                }
                else if (tipo == "inlineStr")
                {
                    texto = c.Element(Ns + "is")?.Value ?? string.Empty;
                }
                else if (v is not null && decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                {
                    numero = d;
                    if (tipo is null)
                    {
                        texto = v;
                    }
                }
                else if (v is not null)
                {
                    texto = v;
                }

                var f = c.Element(Ns + "f");
                hoja.Agregar(referencia, texto, numero, f is not null, f?.Value);
            }
        }

        return hoja;
    }

    private static XDocument Cargar(ZipArchive zip, string entrada)
    {
        var e = zip.GetEntry(entrada) ?? throw new InvalidOperationException($"El xlsx no trae '{entrada}'.");
        using var s = e.Open();
        return XDocument.Load(s);
    }

    private static string[] NombresDeHoja(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        return Cargar(zip, "xl/workbook.xml").Root!.Element(Ns + "sheets")!.Elements(Ns + "sheet")
            .Select(s => s.Attribute("name")!.Value)
            .ToArray();
    }

    /// <summary>Espejo de <c>ExcelDataReaderWorkbookLeafInputReader.ConstruirDetalleR2/R4</c>: primera
    /// fila <c>Vlr Servicio</c>, encabezados por encima, cierre <c>A=Total</c>, filas fantasma
    /// descartadas (sin texto en el label ni valor en ningún componente).</summary>
    private static Fuente LeerFuenteDetalle(int ase, bool esR2, string periodo = "Prueba Julio-2")
    {
        var hoja = Leer(RutaDetalle(ase, esR2, periodo), "Sheet1");
        var colConcepto = esR2 ? "D" : "C";
        var colEncabezados = esR2 ? "E" : "D";

        var primera = Enumerable.Range(2, hoja.MaxRow)
            .First(n => hoja.Texto(colConcepto, n).Equals("Vlr Servicio", StringComparison.OrdinalIgnoreCase));
        Assert.True(primera > 1);

        var filaEncabezados = primera - 1;
        while (filaEncabezados > 1 && hoja.EncabezadosDesde(filaEncabezados, colEncabezados).Length == 0)
        {
            filaEncabezados--;
        }

        var encabezados = hoja.EncabezadosDesde(filaEncabezados, colEncabezados);
        Assert.NotEmpty(encabezados);
        var indices = Enumerable.Range(IndiceDeLetra(colEncabezados), encabezados.Length).ToArray();

        var cierre = Enumerable.Range(primera, hoja.MaxRow - primera + 1)
            .First(n => hoja.Texto("A", n).Equals("Total", StringComparison.OrdinalIgnoreCase));

        var filas = new List<FilaFuente>();
        for (var n = primera; n <= cierre; n++)
        {
            var a = hoja.Texto("A", n);
            var b = hoja.Texto("B", n);
            var c = hoja.Texto("C", n);
            var d = esR2 ? hoja.Texto("D", n) : string.Empty;
            var valores = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < encabezados.Length; i++)
            {
                valores[encabezados[i]] = hoja.NumeroOrNull($"{LetraDe(indices[i])}{n}");
            }

            var hayTexto = !string.IsNullOrWhiteSpace(a) || !string.IsNullOrWhiteSpace(b)
                || !string.IsNullOrWhiteSpace(c) || !string.IsNullOrWhiteSpace(d);
            var hayValor = valores.Values.Any(v => v is not null);
            if (!hayTexto && !hayValor)
            {
                continue;
            }

            filas.Add(new FilaFuente(
                filas.Count, a.Trim(), b.Trim(), c.Trim(), d.Trim(),
                esR2 ? $"{a.Trim()}|{b.Trim()}|{c.Trim()}|{d.Trim()}" : $"{a.Trim()}|{b.Trim()}|{c.Trim()}",
                valores));
        }

        Assert.NotEmpty(filas);
        return new Fuente(filas, encabezados);
    }

    private static string RutaDetalle(int ase, bool esR2, string periodo) =>
        Archivo(CarpetaInsumos(ase, periodo), esR2 ? "RerpoteDetalleSaldosaFavor" : "Reversi");

    private static string RutaSaldosNotas(int ase) =>
        Archivo(CarpetaInsumos(ase, "Prueba Julio-2"), "SaldosaFavorAplicadosPorNotas");

    private static string RutaRetribucionNegativa(int ase) =>
        Archivo(CarpetaInsumos(ase, "Prueba Julio-2"), "Retribuci");

    /// <summary>
    /// Conciliación de julio‑2 resuelta por PREFIJO (no por nombre exacto): el commit <c>88b3acb</c>
    /// reemplazó los insumos y añadió el sufijo de versión (<c>…072026vf.xlsx</c>); congelar el nombre
    /// es la misma clase de defecto que el Plan 59 elimina.
    /// </summary>
    private static string RutaConciliacion(string prefijo)
    {
        var carpeta = Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Insumos", "Conciliaciones");
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta explícita: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .OrderBy(f => f, StringComparer.Ordinal)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta la conciliación '{prefijo}*' en {carpeta}");
        return archivo!;
    }

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    private static string Oficial() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    /// <summary>Plan 48 (H1): base v2 — geometría vigente del contrato R2/R4.</summary>
    private static string V2() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "v2", "Formato Remuneracion 2026xx-x Total vf.xlsx");

    private static string CarpetaInsumos(int ase, string periodo) =>
        Path.Combine(Raiz(), "Docs", periodo, "Insumos", CarpetasAse[ase - 1]);

    private static string Archivo(string carpeta, string prefijo)
    {
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta explícita: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileName(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta la fuente '{prefijo}' en {carpeta}");
        return archivo!;
    }

    private static int IndiceDeLetra(string columna) => columna[0] - 'A';

    private static string LetraDe(int indice) => ((char)('A' + indice)).ToString();
}
