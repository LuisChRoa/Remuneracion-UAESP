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
        private readonly IProcesadorPeriodo _procesadorPeriodo;
        private readonly ArchivoFuenteLocator _archivoFuenteLocator;

        public Form1(
            IProcesadorRemuneracion procesadorRemuneracion,
            IProcesadorPeriodo procesadorPeriodo,
            ArchivoFuenteLocator archivoFuenteLocator)
        {
            _procesadorRemuneracion = procesadorRemuneracion ?? throw new ArgumentNullException(nameof(procesadorRemuneracion));
            _procesadorPeriodo = procesadorPeriodo ?? throw new ArgumentNullException(nameof(procesadorPeriodo));
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
                new Remuneracion.Core.Services.ProcesadorPeriodo(
                    new Remuneracion.Infrastructure.Excel.ExcelDataReaderRecaudoReader(),
                    new Remuneracion.Infrastructure.Excel.ExcelDataReaderWorkbookLeafInputReader(),
                    new Remuneracion.Core.Services.CalculoRemuneracion(),
                    new Remuneracion.Core.Services.ValidadorBasico(),
                    new Remuneracion.Infrastructure.Excel.OpenXmlPlantillaWriter(),
                    new ArchivoFuenteLocator()),
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

        private void chkCincoAse_CheckedChanged(object? sender, EventArgs e)
        {
            // En modo 5 ASE el combo conserva la selección para modo single, pero se deshabilita.
            cmbAse.Enabled = !chkCincoAse.Checked;
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

            var modoCincoAse = chkCincoAse.Checked;
            var maxProgreso = modoCincoAse ? 42 : 8; // hitos × 5 + escritura en modo período

            SetControlesHabilitados(false);
            progressBar.Visible = true;
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Minimum = 0;
            progressBar.Maximum = maxProgreso;
            progressBar.Value = 0;
            toolStripStatusLabel.Text = "Procesando...";

            var periodo = Periodo.Parse(PeriodoSeleccionado);
            var aseSeleccionada = cmbAse.Text;
            var rutaSalida = Path.Combine(txtCarpetaSalida.Text, periodo.NombreArchivo);

            var modoTexto = modoCincoAse ? "5 ASE" : $"ASE {aseSeleccionada}";
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] Iniciando proceso — Período: {PeriodoSeleccionado}, Modo: {modoTexto}{Environment.NewLine}");
            Log.Information("Iniciando proceso de remuneración quincenal. Período: {Periodo}, Modo: {Modo}", PeriodoSeleccionado, modoTexto);

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
                var progreso = new Progress<string>(mensaje =>
                {
                    txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {mensaje}{Environment.NewLine}");
                    Log.Information(mensaje);
                    progressBar.Value = Math.Min(progressBar.Value + 1, progressBar.Maximum);
                });

                if (modoCincoAse)
                {
                    await EjecutarModoCincoAse(periodo, rutaSalida, progreso);
                }
                else
                {
                    await EjecutarModoUnAse(periodo, rutaSalida, progreso);
                }

                progressBar.Value = progressBar.Maximum;
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

        private async Task EjecutarModoUnAse(Periodo periodo, string rutaSalida, IProgress<string> progreso)
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

            var resultadoProceso = await Task.Run(() => _procesadorRemuneracion.Ejecutar(solicitud, progreso));

            var leaf = resultadoProceso.Leaf;
            var valorD9Esperado = leaf.R1.TotalOportunoEsperado;
            var valorF48Esperado = leaf.R1.ExtemporaneoEsperado;
            var valorE41Esperado = leaf.R2.TotalOportunoEsperado;
            var valorD67Esperado = leaf.R4.TotalReversionEsperada;

            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] Resumen final: F46/D9 esperado post-Excel = {valorD9Esperado:0.##}; F48/D47 esperado = {valorF48Esperado:0.##}; E41/D28 esperado = {valorE41Esperado:0.##}; D67/D66 esperado = {valorD67Esperado:0.##}; salida = {rutaSalida}{Environment.NewLine}");
            Log.Information("Resumen final: F46/D9 esperado post-Excel = {D9}; F48/D47 esperado = {D47}; E41/D28 esperado = {D28}; D67/D66 esperado = {D66}; salida = {Salida}",
                valorD9Esperado, valorF48Esperado, valorE41Esperado, valorD67Esperado, rutaSalida);
        }

        private async Task EjecutarModoCincoAse(Periodo periodo, string rutaSalida, IProgress<string> progreso)
        {
            var solicitud = new SolicitudProcesoPeriodo
            {
                Periodo = periodo,
                CarpetaPeriodo = txtCarpetaFuentes.Text,
                RutaPlantilla = txtPlantilla.Text,
                RutaSalida = rutaSalida
            };

            var resultadoProceso = await Task.Run(() => _procesadorPeriodo.Ejecutar(solicitud, progreso));

            // Resumen honesto por ASE: visibles esperados post-Excel (nunca agregados HU-02 como
            // valores CONSOLIDADO) + GranTotal.
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] Resumen por ASE (esperado post-Excel):{Environment.NewLine}");
            Log.Information("Resumen multi-ASE: esperados post-Excel por ASE.");
            foreach (var leaf in resultadoProceso.Leafs.OrderBy(l => l.Ase.Id))
            {
                var linea = $"  ASE {leaf.Ase.Id} {leaf.Ase.NombreCompleto}: TOT_OPT={leaf.R1.TotalOportunoEsperadoPorAse:0.##}; R2={leaf.R2.TotalOportunoEsperado:0.##}; EXTEMP={leaf.R1.ExtemporaneoEsperadoPorAse:0.##}; R4={leaf.R4.TotalReversionEsperada:0.##}";
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {linea}{Environment.NewLine}");
                Log.Information("ASE {AseId}: {Linea}", leaf.Ase.Id, linea);
            }

            // HU-08 (2.2): resumen por empresa de facturación (esperado post-Excel) + Σ vs bloque.
            foreach (var leaf in resultadoProceso.Leafs.OrderBy(l => l.Ase.Id))
            {
                foreach (var conc in leaf.Conciliacion.OrderBy(c => c.Empresa.Id))
                {
                    var lineaEmpresa = $"  ASE {leaf.Ase.Id} · {conc.Empresa.Nombre}: R1={conc.VisibleR1:0.##}; R2={conc.VisibleR2:0.##}; R4={conc.VisibleR4:0.##}";
                    txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {lineaEmpresa}{Environment.NewLine}");
                    Log.Information("ASE {AseId} empresa {Empresa}: {Linea}", leaf.Ase.Id, conc.Empresa.Nombre, lineaEmpresa);
                }
            }

            // HU-08 (2.2): hojas Recaudo * (valores de las conciliaciones por empresa).
            var primerLeaf = resultadoProceso.Leafs.FirstOrDefault();
            if (primerLeaf is not null && primerLeaf.Recaudos.Count > 0)
            {
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] Recaudo por empresa (hojas Recaudo *):{Environment.NewLine}");
                foreach (var recaudo in primerLeaf.Recaudos.OrderBy(r => r.Empresa.Id))
                {
                    var lineaRecaudo = $"  {recaudo.Empresa.Nombre}: OPORTUNO={recaudo.TotalOportuno:0.##}; EXTEMP={recaudo.TotalExtemporaneo:0.##}; TOTAL={recaudo.Total:0.##}";
                    txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {lineaRecaudo}{Environment.NewLine}");
                    Log.Information("Empresa {Empresa}: {Linea}", recaudo.Empresa.Nombre, lineaRecaudo);
                }
            }

            // HU-09 (2.3, §2.6): reporte por banco — por ASE × empresa (4 conceptos + Total,
            // "esperado post-Excel") + C59 + nota 59–80 informativa. Delta mínimo, sin restyle.
            if (resultadoProceso.Leafs.Any(l => l.ReporteBanco is not null))
            {
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] REPORTE RECAUDO x BANCO (esperado post-Excel):{Environment.NewLine}");
                Log.Information("REPORTE RECAUDO x BANCO: esperados post-Excel por ASE y empresa.");
                foreach (var leaf in resultadoProceso.Leafs.OrderBy(l => l.Ase.Id))
                {
                    if (leaf.ReporteBanco is null)
                    {
                        continue;
                    }

                    foreach (var bloque in leaf.ReporteBanco.Ases)
                    {
                        foreach (var empresa in bloque.Empresas.OrderBy(e => e.Empresa))
                        {
                            var lineaBanco = $"  ASE {leaf.Ase.Id} · {empresa.Empresa}: FACT={empresa.AplicadosFacturacion:0.##}; SALDOS={empresa.SaldosFavorGenerados:0.##}; FINANC={empresa.FinanciacionesNuevas:0.##}; ESPEC={empresa.RecibosServEspeciales:0.##}; TOTAL={empresa.Total:0.##}";
                            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {lineaBanco}{Environment.NewLine}");
                            Log.ForContext("Hoja", "REPORTE RECAUDO x BANCO")
                                .Information("ASE {AseId} · {Empresa}: {Linea}", leaf.Ase.Id, empresa.Empresa, lineaBanco);
                        }
                    }
                }

                var quincena = resultadoProceso.Leafs.First(l => l.ReporteBanco is not null).ReporteBanco!.Quincena;
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}]  C59 (quincena) = {quincena}; diferencias filas 59–80 informativas (anulado/reversado misma quincena, esperadas ≠ 0).{Environment.NewLine}");
                Log.ForContext("Hoja", "REPORTE RECAUDO x BANCO")
                    .Information("C59 = {Quincena}; diferencias 59-80 informativas (anulado/reversado).", quincena);
            }

            // HU-10 (2.4, §2.6): balance de subsidios y contribuciones — por ASE (Subsidio E /
            // Contribución D / Total BSC F, "esperado post-Excel") + veredicto D/E citado +
            // nota J9:J13 y K/M calculan por fórmulas (Capa B). Delta mínimo, sin restyle.
            if (resultadoProceso.Leafs.Any(l => l.BalanceSc is not null))
            {
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] BCE SC POR FACT. (esperado post-Excel; asignación D/E = veredicto T0: D←Contribución F-fuente, E←Subsidio E-fuente):{Environment.NewLine}");
                Log.ForContext("Hoja", "BCE SC POR FACT.")
                    .Information("BCE SC POR FACT.: esperados post-Excel por ASE (veredicto D/E T0 hipótesis líder).");
                foreach (var leaf in resultadoProceso.Leafs.OrderBy(l => l.Ase.Id))
                {
                    if (leaf.BalanceSc is null)
                    {
                        continue;
                    }

                    foreach (var bloque in leaf.BalanceSc.Ases)
                    {
                        var lineaBce = $"  ASE {leaf.Ase.Id} {leaf.Ase.NombreCompleto}: CONTRIBUCION(D)={bloque.Contribucion:0.##}; SUBSIDIO(E)={bloque.Subsidio:0.##}; TOTAL BSC(F)={bloque.TotalBsc:0.##}; H≈F por fórmula (DetRetri J{8 + leaf.Ase.Id})";
                        txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {lineaBce}{Environment.NewLine}");
                        Log.ForContext("Hoja", "BCE SC POR FACT.")
                            .Information("ASE {AseId}: {Linea}", leaf.Ase.Id, lineaBce);
                    }
                }

                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}]  CONSOLIDADO J9:J13 y K/M calculan por fórmulas desde BCE F3:F7 (verificar post-Excel en Capa B).{Environment.NewLine}");
                Log.ForContext("Hoja", "BCE SC POR FACT.")
                    .Information("CONSOLIDADO J9:J13 y K/M calculan por fórmulas desde BCE F3:F7 (Capa B).");
            }

                // HU-11 (2.5, §2.6): AJUSTES-SF-T por ASE (saldos-nota / retribución-negativa / total
            // esperado post-Excel) + D85:D89 esperado. Solo en Q2 (leaf.AjustesSfT != null).
            if (resultadoProceso.Leafs.Any(l => l.AjustesSfT is not null))
            {
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] AJUSTES-SF-T (esperado post-Excel; hoja formulada, no escrita):{Environment.NewLine}");
                Log.ForContext("Hoja", "AJUSTES - SF-T")
                    .Information("AJUSTES-SF-T: esperados post-Excel por ASE (Q2).");
                foreach (var leaf in resultadoProceso.Leafs.OrderBy(l => l.Ase.Id))
                {
                    if (leaf.AjustesSfT is null)
                    {
                        continue;
                    }

                    var ajustes = leaf.AjustesSfT;
                    var lineaAjustes = $"  ASE {leaf.Ase.Id} {leaf.Ase.NombreCompleto}: SALDOS-NOTA={ajustes.SaldosNotas.TotalSaldosNotas:0.##}; RETRIBUCION-NEGATIVA={ajustes.RetribucionNegativa.TotalRetribucionNegativa:0.##}; TOTAL AJUSTES(D{84 + leaf.Ase.Id})={ajustes.TotalAjustes:0.##}";
                    txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {lineaAjustes}{Environment.NewLine}");
                    Log.ForContext("Hoja", "AJUSTES - SF-T")
                        .Information("ASE {AseId}: {Linea}", leaf.Ase.Id, lineaAjustes);
                }
            }

            // HU-12 (2.6 ampliada, §2.6): DetRetri-Q2 por ASE (entero ROUND(D104:D108,0) escrito
            // en DetRetri2026072 D9:D13; D14 = total) — solo en Q2 (leaf.DetRetriQ2 != null).
            if (resultadoProceso.Leafs.Any(l => l.DetRetriQ2 is not null))
            {
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] DetRetri Q2 (esperado post-Excel; entero ROUND(D104:D108,0) escrito en valores):{Environment.NewLine}");
                Log.ForContext("Hoja", "DetRetri2026072")
                    .Information("DetRetri Q2: esperados post-Excel por ASE (composición V0.4).");
                foreach (var leaf in resultadoProceso.Leafs.OrderBy(l => l.Ase.Id))
                {
                    if (leaf.DetRetriQ2 is null)
                    {
                        continue;
                    }

                    var detalle = leaf.DetRetriQ2;
                    var lineaDetRetri = $"  ASE {leaf.Ase.Id} {leaf.Ase.NombreCompleto}: D104:D108={detalle.TotalD104:0.##}; DetRetri-D(D{8 + leaf.Ase.Id})={detalle.Detalle:0}";
                    txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {lineaDetRetri}{Environment.NewLine}");
                    Log.ForContext("Hoja", "DetRetri2026072")
                        .Information("ASE {AseId}: {Linea}", leaf.Ase.Id, lineaDetRetri);
                }
            }

            // HU-13 (2.7, §2.6): bloque VALIDACIONES por ASE (veredictos del oráculo de lectura).
            // Solo cuando el procesador trae el lector-oráculo (UI); regresión = lista vacía.
            if (resultadoProceso.Validaciones.Count > 0)
            {
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] VALIDACIONES (oráculo read-only; verificación post-Excel = Capa B manual):{Environment.NewLine}");
                foreach (var linea in resultadoProceso.Validaciones)
                {
                    txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {linea}{Environment.NewLine}");
                    Log.ForContext("Validacion", "cruzada")
                        .Information("{Linea}", linea);
                }
            }

            // GranTotal honesto post-Excel = Σ visibles leaf por ASE (nunca agregados HU-02 como
            // valores CONSOLIDADO; A5).
            var granTotal = resultadoProceso.Leafs.Sum(l =>
                l.R1.TotalOportunoEsperadoPorAse
                + l.R2.TotalOportunoEsperado
                + l.R1.ExtemporaneoEsperadoPorAse
                + l.R4.TotalReversionEsperada
                + (l.AjustesSfT?.TotalAjustes ?? 0m));
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] GranTotal CONSOLIDADO (Σ visibles post-Excel) = {granTotal:0.##}; salida = {rutaSalida}{Environment.NewLine}");
            Log.Information("GranTotal CONSOLIDADO (Σ visibles post-Excel) = {GranTotal}; salida = {Salida}", granTotal, rutaSalida);
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
            cmbAse.Enabled = habilitados && !chkCincoAse.Checked;
            chkCincoAse.Enabled = habilitados;
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