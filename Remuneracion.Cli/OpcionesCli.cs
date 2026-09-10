using Remuneracion.Core.Models;

namespace Remuneracion.Cli;

/// <summary>
/// HU-15 (3.3, D2): parseo MANUAL de argumentos del CLI (sin System.CommandLine; 0 paquetes
/// nuevos). Puro y testeable in-memory: no toca I/O, Serilog ni <c>Environment.Exit</c>.
/// Errores de uso → <see cref="UsoCliException"/> (stderr + ayuda + salida 4, G6); ningún
/// código de error nuevo: el contrato <see cref="Remuneracion.Core.Errors.CodigosSalida"/> 0-5
/// queda intacto.
/// </summary>
public sealed class OpcionesCli
{
    /// <summary>Código completo del período resuelto (AAAAMMQ), p. ej. "2026071".</summary>
    public string PeriodoCodigoCompleto { get; }

    /// <summary>Período resuelto por <see cref="Periodo.Parse"/>.</summary>
    public Periodo Periodo { get; }

    /// <summary>Carpeta del período que contiene las 5 carpetas de ASE.</summary>
    public string Carpeta { get; }

    /// <summary>Ruta de la plantilla de origen.</summary>
    public string Plantilla { get; }

    /// <summary>Carpeta de salida (G5): el archivo se nombra por <see cref="Periodo.NombreArchivo"/>.</summary>
    public string Salida { get; }

    /// <summary>ASE a procesar en modo 1-ASE (1..5); null = modo 5 ASE.</summary>
    public int? AseId { get; }

    /// <summary>Modo 5 ASE por defecto (G6); falso solo si se indicó <c>--ase N</c>.</summary>
    public bool CincoAse => AseId is null;

    /// <summary>Permite sobrescribir la salida si ya existe.</summary>
    public bool Sobrescribir { get; }

    /// <summary><c>--help</c>/<c>-h</c>: imprime uso a stdout e ignora el resto (salida 0).</summary>
    public bool Ayuda { get; }

    public string ModoTexto => CincoAse ? "5 ASE" : $"ASE {AseId}";

    private OpcionesCli(
        string periodoCodigoCompleto,
        Periodo periodo,
        string carpeta,
        string plantilla,
        string salida,
        int? aseId,
        bool sobrescribir,
        bool ayuda)
    {
        PeriodoCodigoCompleto = periodoCodigoCompleto;
        Periodo = periodo;
        Carpeta = carpeta;
        Plantilla = plantilla;
        Salida = salida;
        AseId = aseId;
        Sobrescribir = sobrescribir;
        Ayuda = ayuda;
    }

    /// <summary>
    /// Parsea los argumentos según la tabla del contrato (HU-15 §2.1). Lanza
    /// <see cref="UsoCliException"/> ante cualquier violación de uso.
    /// </summary>
    public static OpcionesCli Parse(string[] args)
    {
        ArgumentNullException.ThrowIfNull(args);

        string? periodoRaw = null;
        int? quincena = null;
        string? carpeta = null;
        string? plantilla = null;
        string? salida = null;
        int? aseId = null;
        var cincoAseExplicito = false;
        var sobrescribir = false;

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            switch (arg)
            {
                case "--help" or "-h":
                    // HU-15 §2.1: --help imprime uso a stdout e IGNORA el resto (salida 0).
                    // Early-return aquí: la ayuda presente gana aunque haya args inválidos
                    // tras ella (p. ej. `--help --ase 9` → 0, no 4). W-1.
                    return new OpcionesCli(string.Empty, new Periodo(), string.Empty, string.Empty, string.Empty, null, false, true);
                case "--periodo":
                    periodoRaw = ValorSiguiente(args, ref i, "--periodo");
                    break;
                case "--quincena":
                    if (!int.TryParse(ValorSiguiente(args, ref i, "--quincena"), out var q) || q is not 1 and not 2)
                    {
                        throw new UsoCliException($"--quincena debe ser 1 o 2.");
                    }

                    quincena = q;
                    break;
                case "--carpeta":
                    carpeta = ValorSiguiente(args, ref i, "--carpeta");
                    break;
                case "--plantilla":
                    plantilla = ValorSiguiente(args, ref i, "--plantilla");
                    break;
                case "--salida":
                    salida = ValorSiguiente(args, ref i, "--salida");
                    break;
                case "--ase":
                    if (cincoAseExplicito)
                    {
                        throw new UsoCliException("--ase no puede combinarse con --cinco-ase (son mutuamente excluyentes).");
                    }

                    var valorAse = ValorSiguiente(args, ref i, "--ase");
                    if (!int.TryParse(valorAse, out var idAse) || idAse is < 1 or > 5)
                    {
                        throw new UsoCliException($"--ase debe ser un número entre 1 y 5; recibió '{valorAse}'.");
                    }

                    aseId = idAse;
                    break;
                case "--cinco-ase":
                    if (aseId is not null)
                    {
                        throw new UsoCliException("--cinco-ase no puede combinarse con --ase (son mutuamente excluyentes).");
                    }

                    cincoAseExplicito = true;
                    break;
                case "--sobrescribir":
                    sobrescribir = true;
                    break;
                default:
                    throw new UsoCliException($"Flag desconocido: '{arg}'.");
            }
        }

        if (periodoRaw is null)
        {
            throw new UsoCliException("--periodo es obligatorio (AAAAMM o AAAAMMQ).");
        }

        if (!EsCodigoPeriodoValido(periodoRaw))
        {
            throw new UsoCliException($"El período '{periodoRaw}' no es válido: use AAAAMM (6 dígitos) o AAAAMMQ (7 dígitos con quincena 1 o 2).");
        }

        var quincenaDelPeriodo = periodoRaw.Length == 7 ? periodoRaw[^1] - '0' : (int?)null;
        if (quincenaDelPeriodo is not null && quincena is not null && quincenaDelPeriodo != quincena)
        {
            throw new UsoCliException($"La quincena de --quincena ({quincena}) difiere de la del período ({quincenaDelPeriodo}).");
        }

        var quincenaFinal = quincenaDelPeriodo ?? quincena;
        if (quincenaFinal is null)
        {
            throw new UsoCliException($"El período '{periodoRaw}' no trae quincena; use --quincena 1|2.");
        }

        if (carpeta is null)
        {
            throw new UsoCliException("--carpeta es obligatoria (carpeta del período con las 5 carpetas de ASE).");
        }

        if (plantilla is null)
        {
            throw new UsoCliException("--plantilla es obligatoria (plantilla de origen).");
        }

        if (salida is null)
        {
            throw new UsoCliException("--salida es obligatoria (carpeta de salida).");
        }

        var codigoCompleto = periodoRaw.Length == 7 ? periodoRaw : periodoRaw + quincenaFinal.Value;
        Periodo periodo;
        try
        {
            periodo = Periodo.Parse(codigoCompleto);
        }
        catch (ArgumentException ex)
        {
            throw new UsoCliException(ex.Message);
        }

        return new OpcionesCli(codigoCompleto, periodo, carpeta, plantilla, salida, aseId, sobrescribir, false);
    }

    /// <summary>
    /// Texto de uso/ayuda del CLI (stdout para --help; stderr + salida 4 para errores de uso).
    /// </summary>
    public static string Uso() =>
        """
        Uso: Remuneracion.Cli --periodo <AAAAMM|AAAAMMQ> --carpeta <dir-periodo> --plantilla <xlsx> --salida <dir>
                             [--ase N | --cinco-ase] [--quincena 1|2] [--sobrescribir] [--help]

          --periodo AAAAMM[Q]  Período (p. ej. 2026071). Si usa AAAAMM, --quincena es obligatoria.
          --quincena 1|2       Quincena (override/complemento del período).
          --carpeta <dir>      Carpeta del período que contiene las 5 carpetas de ASE.
          --plantilla <xlsx>   Plantilla de origen (se copia; jamás se modifica).
          --salida <dir>       Carpeta de salida (se crea si falta; el archivo se nombra por el período).
          --ase N              Procesa un solo ASE (1..5). Excluye --cinco-ase.
          --cinco-ase          Procesa los 5 ASE (predeterminado).
          --sobrescribir       Sobrescribe la salida si ya existe.
          --help, -h           Muestra esta ayuda y sale con 0.

        Códigos de salida (contrato HU-14): 0 OK · 1 validación · 2 fuente/plantilla ·
        3 escritura · 4 inesperado · 5 cancelado.
        """;

    private static bool EsCodigoPeriodoValido(string periodo)
    {
        if (periodo.Length is not 6 and not 7 || !periodo.All(char.IsAsciiDigit))
        {
            return false;
        }

        if (periodo.Length == 7 && periodo[^1] is not '1' and not '2')
        {
            return false;
        }

        return true;
    }

    private static string ValorSiguiente(string[] args, ref int indice, string flag)
    {
        if (indice + 1 >= args.Length)
        {
            throw new UsoCliException($"{flag} requiere un valor.");
        }

        indice++;
        return args[indice];
    }
}

/// <summary>
/// Error de USO del CLI (G6): flag desconocido, período malformado, exclusión violada, flag
/// requerido ausente. El ejecutor lo traduce a stderr + ayuda + salida 4. NO es un error de
/// dominio: no crea ningún código del catálogo.
/// </summary>
public sealed class UsoCliException : Exception
{
    public UsoCliException(string message) : base(message)
    {
    }
}
