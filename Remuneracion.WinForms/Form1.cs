using System.IO;
using Serilog;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.FileSystem;

namespace Remuneracion.WinForms
{
    public partial class Form1 : Form
    {
        private static readonly string[] MesesEspanol =
        [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ];

        private readonly IProcesadorRemuneracion _procesadorRemuneracion;
        private readonly ArchivoFuenteLocator _archivoFuenteLocator;

        public Form1(IProcesadorRemuneracion procesadorRemuneracion, ArchivoFuenteLocator archivoFuenteLocator)
        {
            _procesadorRemuneracion = procesadorRemuneracion ?? throw new ArgumentNullException(nameof(procesadorRemuneracion));
            _archivoFuenteLocator = archivoFuenteLocator ?? throw new ArgumentNullException(nameof(archivoFuenteLocator));

            InitializeComponent();
            ConfigurarSerilog();
            InicializarPeriodo();
            InicializarAse();
        }

        public Form1()
            : this(
                new Remuneracion.Core.Services.ProcesadorRemuneracion(
                    new Remuneracion.Infrastructure.Excel.ExcelDataReaderRecaudoReader(),
                    new Remuneracion.Infrastructure.Excel.ExcelDataReaderWorkbookLeafInputReader(),
                    new Remuneracion.Core.Services.CalculoRemuneracion(),
                    new Remuneracion.Core.Services.ValidadorBasico(),
                    new Remuneracion.Infrastructure.Excel.OpenXmlPlantillaWriter()),
                new ArchivoFuenteLocator())
        {
        }

        public string PeriodoSeleccionado
        {
            get
            {
                if (cmbAnio.SelectedItem == null || cmbMes.SelectedIndex < 0 || cmbQuincena.SelectedIndex < 0)
                    return string.Empty;

                string anio = cmbAnio.Text;
                string mes = (cmbMes.SelectedIndex + 1).ToString("D2");
                string quincena = (cmbQuincena.SelectedIndex + 1).ToString();
                return $"{anio}{mes}{quincena}";
            }
        }

        private void InicializarPeriodo()
        {
            int anioActual = DateTime.Now.Year;

            cmbAnio.BeginUpdate();
            for (int a = anioActual - 5; a <= anioActual + 2; a++)
            {
                cmbAnio.Items.Add(a.ToString());
            }
            cmbAnio.SelectedItem = anioActual.ToString();
            cmbAnio.EndUpdate();

            cmbMes.BeginUpdate();
            cmbMes.Items.AddRange(MesesEspanol);
            cmbMes.SelectedIndex = DateTime.Now.Month - 1;
            cmbMes.EndUpdate();

            cmbQuincena.Items.AddRange(["1.ª Quincena", "2.ª Quincena"]);
            cmbQuincena.SelectedIndex = DateTime.Now.Day <= 15 ? 0 : 1;
        }

        private void InicializarAse()
        {
            cmbAse.BeginUpdate();
            foreach (string prefijo in CarpetasAse.Prefijos)
            {
                int guion = prefijo.IndexOf('-');
                if (guion > 0)
                {
                    string numero = prefijo[..guion].Trim();
                    string nombre = prefijo[(guion + 1)..].Trim();
                    cmbAse.Items.Add($"{numero} - {nombre}");
                }
                else
                {
                    cmbAse.Items.Add(prefijo);
                }
            }
            cmbAse.SelectedIndex = 0;
            cmbAse.EndUpdate();
        }

        private void btnSeleccionarCarpeta_Click(object? sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                txtCarpetaFuentes.Text = folderBrowserDialog.SelectedPath;
            }
        }

        private void btnSeleccionarPlantilla_Click(object? sender, EventArgs e)
        {
            if (openFileDialogPlantilla.ShowDialog() == DialogResult.OK)
            {
                txtPlantilla.Text = openFileDialogPlantilla.FileName;
            }
        }

        private void btnSeleccionarSalida_Click(object? sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                txtCarpetaSalida.Text = folderBrowserDialog.SelectedPath;
            }
        }

        private async void btnEjecutar_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCarpetaFuentes.Text) ||
                string.IsNullOrWhiteSpace(txtPlantilla.Text) ||
                string.IsNullOrWhiteSpace(txtCarpetaSalida.Text))
            {
                MessageBox.Show(
                    "Debe seleccionar la carpeta de fuentes, la plantilla y la carpeta de salida antes de ejecutar.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            SetControlesHabilitados(false);
            progressBar.Visible = true;
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Minimum = 0;
            progressBar.Maximum = 8;
            progressBar.Value = 0;
            toolStripStatusLabel.Text = "Procesando...";

            var periodo = Periodo.Parse(PeriodoSeleccionado);
            var aseSeleccionada = cmbAse.Text;
            var rutaSalida = Path.Combine(txtCarpetaSalida.Text, periodo.NombreArchivo);

            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] Iniciando proceso — Período: {PeriodoSeleccionado}, ASE: {aseSeleccionada}{Environment.NewLine}");
            Log.Information("Iniciando proceso de remuneración quincenal. Período: {Periodo}, ASE: {Ase}", PeriodoSeleccionado, aseSeleccionada);

            if (string.Equals(Path.GetFullPath(txtPlantilla.Text), Path.GetFullPath(rutaSalida), StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("La ruta de salida debe ser distinta a la plantilla original.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetControlesHabilitados(true);
                progressBar.Visible = false;
                return;
            }

            if (File.Exists(rutaSalida))
            {
                var resultado = MessageBox.Show(
                    $"El archivo '{Path.GetFileName(rutaSalida)}' ya existe. ¿Desea sobrescribirlo?",
                    "Archivo existente",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (resultado != DialogResult.Yes)
                {
                    txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] Proceso cancelado por decisión del usuario. Archivo de salida ya existe.{Environment.NewLine}");
                    Log.Information("Proceso cancelado: salida ya existe y no se acepta sobreescritura.");
                    SetControlesHabilitados(true);
                    progressBar.Visible = false;
                    return;
                }
            }

            try
            {
                var ase = ParseAse(cmbAse.Text);
                var carpetaAse = ObtenerCarpetaAse(ase.Id);

                var solicitud = new SolicitudProcesoAse
                {
                    Ase = ase,
                    Periodo = periodo,
                    RutaR1 = _archivoFuenteLocator.BuscarArchivo(carpetaAse, "Recaudoporcomponente") ?? throw new ArchivoFuenteNoEncontradoException($"No se encontró R1 en {carpetaAse}."),
                    RutaR2 = _archivoFuenteLocator.BuscarArchivo(carpetaAse, "RerpoteDetalleSaldosaFavor") ?? throw new ArchivoFuenteNoEncontradoException($"No se encontró R2 en {carpetaAse}."),
                    RutaR4 = _archivoFuenteLocator.BuscarArchivo(carpetaAse, "ReversiónPorComponente") ?? _archivoFuenteLocator.BuscarArchivo(carpetaAse, "ReversionPorComponente") ?? throw new ArchivoFuenteNoEncontradoException($"No se encontró R4 en {carpetaAse}."),
                    RutaPlantilla = txtPlantilla.Text,
                    RutaSalida = rutaSalida
                };

                var progreso = new Progress<string>(mensaje =>
                {
                    txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {mensaje}{Environment.NewLine}");
                    Log.Information(mensaje);
                    progressBar.Value = Math.Min(progressBar.Value + 1, progressBar.Maximum);
                });

                var resultadoProceso = await Task.Run(() => _procesadorRemuneracion.Ejecutar(solicitud, progreso));
                progressBar.Value = progressBar.Maximum;

                var leaf = resultadoProceso.Leaf;
                var valorD9Esperado = leaf.R1.TotalOportunoEsperado;
                var valorF48Esperado = leaf.R1.ExtemporaneoEsperado;
                var valorE41Esperado = leaf.R2.TotalOportunoEsperado;
                var valorD67Esperado = leaf.R4.TotalReversionEsperada;

                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] Resumen final: F46/D9 esperado post-Excel = {valorD9Esperado:0.##}; F48/D47 esperado = {valorF48Esperado:0.##}; E41/D28 esperado = {valorE41Esperado:0.##}; D67/D66 esperado = {valorD67Esperado:0.##}; salida = {rutaSalida}{Environment.NewLine}");
                Log.Information("Resumen final: F46/D9 esperado post-Excel = {D9}; F48/D47 esperado = {D47}; E41/D28 esperado = {D28}; D67/D66 esperado = {D66}; salida = {Salida}",
                    valorD9Esperado, valorF48Esperado, valorE41Esperado, valorD67Esperado, rutaSalida);

                toolStripStatusLabel.Text = "Completado";
            }
            catch (Exception ex)
            {
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] ERROR: {ex.GetType().Name} — {ex.Message}{Environment.NewLine}");
                Log.Error(ex, "Error en la ejecución del proceso.");
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                toolStripStatusLabel.Text = "Error";
            }
            finally
            {
                progressBar.Visible = false;
                SetControlesHabilitados(true);
            }
        }

        private string ObtenerCarpetaAse(int idAse)
        {
            var carpetaPeriodo = txtCarpetaFuentes.Text;
            var carpeta = _archivoFuenteLocator.ObtenerCarpetasAse(carpetaPeriodo)
                .FirstOrDefault(c => Path.GetFileName(c).StartsWith(idAse.ToString(), StringComparison.OrdinalIgnoreCase));

            if (string.IsNullOrWhiteSpace(carpeta))
            {
                throw new ArchivoFuenteNoEncontradoException($"No se encontró la carpeta del ASE {idAse} en '{carpetaPeriodo}'.");
            }

            return carpeta;
        }

        private static Ase ParseAse(string texto)
        {
            var numero = texto.Split('-')[0].Trim();
            var nombre = texto.Contains('-') ? texto[(texto.IndexOf('-') + 1)..].Trim() : texto.Trim();
            return new Ase
            {
                Id = int.Parse(numero),
                NombreCompleto = nombre,
                NombreCorto = nombre.ToUpperInvariant(),
                NumeroCarpeta = int.Parse(numero)
            };
        }

        private void SetControlesHabilitados(bool habilitados)
        {
            cmbAnio.Enabled = habilitados;
            cmbMes.Enabled = habilitados;
            cmbQuincena.Enabled = habilitados;
            txtCarpetaFuentes.Enabled = habilitados;
            btnSeleccionarCarpeta.Enabled = habilitados;
            txtPlantilla.Enabled = habilitados;
            btnSeleccionarPlantilla.Enabled = habilitados;
            txtCarpetaSalida.Enabled = habilitados;
            btnSeleccionarSalida.Enabled = habilitados;
            cmbAse.Enabled = habilitados;
            btnEjecutar.Enabled = habilitados;
        }

        private void ConfigurarSerilog()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.File("remuneracion_log.txt")
                .CreateLogger();
        }
    }
}
