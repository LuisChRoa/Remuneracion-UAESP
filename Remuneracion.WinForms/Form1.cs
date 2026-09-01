using Serilog;
using Remuneracion.Core.Constants;

namespace Remuneracion.WinForms
{
    public partial class Form1 : Form
    {
        private static readonly string[] MesesEspanol =
        [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ];

        public Form1()
        {
            InitializeComponent();
            ConfigurarSerilog();
            InicializarPeriodo();
            InicializarAse();
        }

        /// <summary>
        /// Construye el código de período AAAAMM# a partir de los 3 ComboBox.
        /// Ejemplo: 2026 + Julio + 1.ª → "2026071".
        /// </summary>
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

            // Años: 5 hacia atrás, 2 hacia adelante
            cmbAnio.BeginUpdate();
            for (int a = anioActual - 5; a <= anioActual + 2; a++)
            {
                cmbAnio.Items.Add(a.ToString());
            }
            cmbAnio.SelectedItem = anioActual.ToString();
            cmbAnio.EndUpdate();

            // Meses en español
            cmbMes.BeginUpdate();
            cmbMes.Items.AddRange(MesesEspanol);
            cmbMes.SelectedIndex = DateTime.Now.Month - 1;
            cmbMes.EndUpdate();

            // Quincena: detectar por día del mes (1-15 = 1.ª)
            cmbQuincena.Items.AddRange(["1.ª Quincena", "2.ª Quincena"]);
            cmbQuincena.SelectedIndex = DateTime.Now.Day <= 15 ? 0 : 1;
        }

        private void InicializarAse()
        {
            cmbAse.BeginUpdate();
            foreach (string prefijo in CarpetasAse.Prefijos)
            {
                // Formato de salida: "1 - Promoambiental" a partir de "1-Promoambiental"
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

        private async void btnEjecutar_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCarpetaFuentes.Text) ||
                string.IsNullOrWhiteSpace(txtPlantilla.Text))
            {
                MessageBox.Show(
                    "Debe seleccionar la carpeta de fuentes y la plantilla antes de ejecutar.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            // Deshabilitar controles durante la ejecución
            SetControlesHabilitados(false);
            progressBar.Visible = true;
            toolStripStatusLabel.Text = "Procesando...";

            txtLog.AppendText(
                $"[{DateTime.Now:HH:mm:ss}] Iniciando proceso \u2014 Período: {PeriodoSeleccionado}, ASE: {cmbAse.Text}"
                + Environment.NewLine);
            Log.Information(
                "Iniciando proceso de remuneración quincenal (stub). Período: {Periodo}, ASE: {Ase}",
                PeriodoSeleccionado, cmbAse.Text);

            txtLog.AppendText(
                $"[{DateTime.Now:HH:mm:ss}] Los readers/writers de Excel son stubs (NotImplementedException)."
                + Environment.NewLine);
            Log.Information("Readers/writers pendientes de implementación con datos reales.");

            // Simular trabajo (stub)
            await Task.Delay(2000);

            progressBar.Visible = false;
            toolStripStatusLabel.Text = "Completado (stub)";
            txtLog.AppendText(
                $"[{DateTime.Now:HH:mm:ss}] Proceso finalizado (stub)."
                + Environment.NewLine);
            Log.Information("Proceso finalizado (stub).");
            SetControlesHabilitados(true);
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
