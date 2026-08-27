using Serilog;

namespace Remuneracion.WinForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
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

        private void btnEjecutar_Click(object? sender, EventArgs e)
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

            ConfigurarSerilog();

            progressBar.Visible = true;
            lblEstado.Text = "Procesando...";
            txtLog.AppendText("Iniciando proceso..." + Environment.NewLine);
            Log.Information("Iniciando proceso de remuneración quincenal (stub).");

            txtLog.AppendText(
                "Los readers/writers de Excel son stubs (NotImplementedException). " +
                "Se requieren archivos de plantilla y fuentes de ejemplo para implementar la lógica real." +
                Environment.NewLine);
            Log.Information("Readers/writers pendientes de implementación con datos reales.");

            lblEstado.Text = "Completado (stub)";
            progressBar.Visible = false;
        }

        private static void ConfigurarSerilog()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.File("remuneracion_log.txt")
                .CreateLogger();
        }
    }
}
